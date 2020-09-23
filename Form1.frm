VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Print LynxGrid"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportCSV 
      Caption         =   "Export CSV"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "Print Preview"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin Project1.LynxGrid LynxGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebPreview 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   975
      ExtentX         =   1720
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function getSaveCSVFileName() As String

    With CommonDialog1
        .FileName = vbNullString
        .DialogTitle = "Save File as..."
        .InitDir = App.Path ' 1.88
        .Flags = cdlOFNHideReadOnly
        .Filter = "Comma Delimited File (*.csv)|*.csv;"
        .CancelError = False
        .ShowSave
        getSaveCSVFileName = .FileName
    End With 'FRMMASTER.MAPDIAG

End Function

Private Sub cmdExportCSV_Click()

Dim FullFileName As String
Dim newLynxPrint As New clsLynxPrint

    FullFileName = getSaveCSVFileName
    If LenB(FullFileName) <> 0 Then
        newLynxPrint.LynxGridExportToCSV LynxGrid1, FullFileName
    End If

End Sub

Private Sub cmdPrint_Click()

Dim newLynxPrint As New clsLynxPrint
    newLynxPrint.DocTitle = "Hello"
    newLynxPrint.PrintLynxGrid LynxGrid1, WebPreview, StraightToPrinter

End Sub

Private Sub cmdPrintPreview_Click()

Dim newLynxPrint As New clsLynxPrint
    newLynxPrint.DocTitle = "Hello"
    newLynxPrint.PrintLynxGrid LynxGrid1, WebPreview, PrintPreview

End Sub

Private Sub Form_Load()

    WebPreview.Move 0, 0, 0, 0
    WebPreview.Visible = False
    popLynxGrid1

End Sub

Private Sub popLynxGrid1()

Dim lnRowCounter As Long

    With LynxGrid1
        .Redraw = False
        .AddColumn "Col 1", 1500
        .AddColumn "Col 2", 1500
        .AddColumn "Col 3", 1500
        .AddColumn "Col 4", 1500
        For lnRowCounter = 0 To 100
            .AddItem "A" & lnRowCounter & vbTab & _
                     "B" & lnRowCounter & vbTab & _
                     "C" & lnRowCounter & vbTab & _
                     "D" & lnRowCounter
        Next lnRowCounter
        .Redraw = True
    End With

End Sub
