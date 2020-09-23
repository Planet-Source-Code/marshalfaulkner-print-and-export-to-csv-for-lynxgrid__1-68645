VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLynxPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'printing portions modified from Hawke's "Print preview using xml"
'html generation modified from Eoin Armstrong's "MSFlexgrid to HTML"
'To use with another grid simply modify the GridToArray Procedure

Public Enum PrintPrompt
    PrintPreview = 0
    StraightToPrinter = 1
End Enum

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long

Private Const navNoHistory As Integer = 2
Private Const navNoWriteToCache As Integer = 8
Private m_TempPath As String
Private m_Title As String
Private HeaderArray() As String
Private BodyArray() As String
Private m_VisibleColumnsOnly As Boolean

Public Sub PrintLynxGrid(LynxGrid1 As LynxGrid, _
                         WebPreview As WebBrowser, _
                         PreviewType As PrintPrompt)

    m_TempPath = GetTempDirectory & CreateGUID & ".html"
    GridToArray LynxGrid1, m_VisibleColumnsOnly
    ArraysToHTML
    ProcessPrinting WebPreview, PreviewType

End Sub

Public Property Let DocTitle(strTitle As String)
    m_Title = strTitle
End Property

Public Property Get DocTitle() As String
    DocTitle = m_Title
End Property

Private Sub ArraysToHTML()

Dim strm As TextStream
Dim lnRowCounter As Long
Dim lnColCounter As Long
Dim fso As New FileSystemObject
    
    Set strm = fso.CreateTextFile(m_TempPath, True)
    strm.Write ("<html><head>" & vbNewLine)
    
    strm.Write ("<title>" & m_Title & "</title>" & vbNewLine)
    strm.Write ("<table border=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine)
    strm.Write ("<tr>" & vbNewLine)
    
    'write column headings
    For lnColCounter = 0 To UBound(HeaderArray)
        strm.Write ("   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;" & HeaderArray(lnColCounter) & "&nbsp;&nbsp;</b></td>" & vbNewLine)
    Next lnColCounter
    strm.Write ("</tr>" & vbNewLine)
    
    'write body
    For lnRowCounter = 0 To UBound(BodyArray, 2)
        strm.Write ("<tr>" & vbNewLine)
        For lnColCounter = 0 To UBound(BodyArray, 1)
            strm.Write ("   <td align=" & Chr(34) & "left" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & ">&nbsp;&nbsp;" & BodyArray(lnColCounter, lnRowCounter) & "&nbsp;&nbsp;</td>" & vbNewLine)
        Next lnColCounter
        strm.Write ("</tr>" & vbNewLine)
    Next lnRowCounter
    
    'closing off tags
    strm.Write ("</table></head></html>")

End Sub

Private Sub GridToArray(LynxGrid1 As LynxGrid, _
                        VisibleColumnsOnly As Boolean)

Dim lnRowCounter As Long
Dim lnColCounter As Long
Dim numVisibleCols As Long
Dim currCol As Long

    With LynxGrid1
        'get number of columns for Array Size
        For lnColCounter = 0 To .Cols
            'visible?
            If (.ColVisible(lnColCounter) And .ColWidth(lnColCounter) > 0) Or Not VisibleColumnsOnly Then
                numVisibleCols = numVisibleCols + 1
            End If
        Next lnColCounter
        ReDim HeaderArray(numVisibleCols - 1)
        ReDim BodyArray(numVisibleCols - 1, 0)
        'start headings section -----------
        For lnColCounter = 0 To .Cols
            'visible?
            If (.ColVisible(lnColCounter) And .ColWidth(lnColCounter) > 0) Or Not VisibleColumnsOnly Then
                HeaderArray(currCol) = .ColHeading(lnColCounter) & ""
                currCol = currCol + 1
            End If
        Next lnColCounter
        'end headings section -------------
        
        'start body section ----------
        For lnRowCounter = 0 To .ItemCount - 1
            currCol = 0
            For lnColCounter = 0 To .Cols
                'visible?
                If (.ColVisible(lnColCounter) And .ColWidth(lnColCounter) > 0) Or Not VisibleColumnsOnly Then
                    BodyArray(currCol, UBound(BodyArray, 2)) = .CellText(lnRowCounter, lnColCounter) & ""
                    currCol = currCol + 1
                End If
            Next lnColCounter
            ReDim Preserve BodyArray(UBound(BodyArray, 1), UBound(BodyArray, 2) + 1)
        Next lnRowCounter
        ReDim Preserve BodyArray(UBound(BodyArray, 1), UBound(BodyArray, 2) - 1)
        'end body section -----------
    End With

End Sub

Private Sub ProcessPrinting(WebPreview As WebBrowser, PreviewType As PrintPrompt)

    With WebPreview
        .Visible = True
        'navigate to page
        .Navigate2 m_TempPath, navNoHistory & navNoWriteToCache
        'loop until page is finished loading
        While .ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Wend
        'print
        If PreviewType = PrintPreview Then
            .ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        ElseIf PreviewType = StraightToPrinter Then
            .ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        End If
        .Visible = False
    End With

End Sub

Private Function CreateGUID() As String

    Dim bytID(0 To 15) As Byte
    Dim lngCount As Long
    
    If CoCreateGuid(bytID(0)) = 0 Then
        For lngCount = 0 To 15
            CreateGUID = CreateGUID + IIf(bytID(lngCount) < 16, "0", "") + Hex$(bytID(lngCount))
        Next
        
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    End If
    
End Function

Private Function GetTempDirectory() As String
    
    Dim strTemp As String
    Dim strUserName As String
        
    strTemp = String(100, Chr$(0))  'Create a buffer
    GetTempPath 100, strTemp
    strTemp = Trim(Left$(strTemp, InStr(strTemp, Chr$(0)) - 1))
    
    If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    GetTempDirectory = strTemp
    
End Function

Public Sub LynxGridExportToCSV(LynxGrid1 As LynxGrid, _
                               filePath As String, _
                               Optional strHeader As String = vbNullString, _
                               Optional exportColHeaders As Boolean = True, _
                               Optional VisibleColumnsOnly As Boolean = True)

Dim strm         As TextStream
Dim lnRowCounter As Long
Dim lnColCounter As Long
Dim fso          As New FileSystemObject
Dim firstColHit As Boolean

    GridToArray LynxGrid1, VisibleColumnsOnly
    
    Set strm = fso.CreateTextFile(filePath, True)
    With LynxGrid1
        'start document heading section ---------
        If LenB(strHeader) <> 0 Then
            strm.Write (Chr(34) & strHeader & Chr(34))
            strm.Write (vbNewLine)
        End If
        'end document heading section ---------
        
        'start headings section -----------
        If exportColHeaders Then
            For lnColCounter = 0 To UBound(HeaderArray)
                If firstColHit Then
                    strm.Write (",")
                End If
                'write
                strm.Write (Chr(34) & HeaderArray(lnColCounter) & "") & Chr(34)
                firstColHit = True
            Next lnColCounter
            strm.Write (vbNewLine)
        End If
        'end headings section -------------
        
        'start body section ----------
        For lnRowCounter = 0 To UBound(BodyArray, 2)
            firstColHit = False
            For lnColCounter = 0 To UBound(BodyArray, 1)
                If firstColHit Then
                    strm.Write (",")
                End If
                'write
                strm.Write (Chr(34) & BodyArray(lnColCounter, lnRowCounter) & "") & Chr(34)
                firstColHit = True
            Next lnColCounter
            strm.Write (vbNewLine)
        Next lnRowCounter
        'end body section -----------
    End With

End Sub

Private Sub Class_Terminate()

    'delete file after finished
    If LenB(m_TempPath) <> 0 Then
        Kill m_TempPath
    End If

End Sub
