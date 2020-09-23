VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test form"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrintBackground 
      Caption         =   "Print Background"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   3915
      Begin VB.Label Label1 
         Caption         =   "Note that WebBrowser control named wbReport is placed on this form (outside the visible region)."
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4140
      Top             =   60
   End
   Begin SHDocVwCtl.WebBrowser wbReport 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   -600
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   979
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show report"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HTML Reports 1.01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Created by Tomasz Puwalski"
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   2040
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
' HTML Reports - Test form
' Created by Tomasz Puwalski (pvl@cps.pl)
' Version 1.01 (09/17/2003)
'---------------------------------------------
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Constants for ShowWindow
Const SW_NORMAL = 1
Const SW_MAXIMIZE = 3
Const SW_MINIMIZE = 6

Const OLECMDID_PRINTPREVIEW = 7

Private blnReportCreated As Boolean
Private lngHWnd As Long
Private strPrintBackKey As String
Private strPrintBackChecked As String
Private strPrintBackUnchecked As String


Private Sub chkPrintBackground_Click()
  WriteKey strPrintBackKey, IIf(chkPrintBackground.Value = Checked, strPrintBackChecked, strPrintBackUnchecked)
End Sub

Private Sub Command1_Click()
  Dim objReport As New clsReport
  Dim objTableConfig As New clsTableConfig
  Dim dbConn As New Connection
  Dim RS As Recordset
  Dim strSeparator As String
  
  dbConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\wta.mdb"
  Set RS = dbConn.Execute("SELECT * FROM wta")
  
  With objTableConfig
    .Border = 1
    .CellPadding = 1
    .Class = "DataTable"
    .Width = "90%"
    .AddHeaderForField "NextOff", "Next Off"
    .AddAlignForField "Rank", [Right Justify]
    .AddAlignForField "Trend", [Right Justify]
    .AddAlignForField "Points", [Right Justify]
    .AddAlignForField "Tournament", [Right Justify]
    .AddAlignForField "NextOff", [Right Justify]
    .AddWidthForField "Rank", "7%"
    .AddWidthForField "Trend", "7%"
    .AddWidthForField "Name", "30%"
    .AddWidthForField "Country", "24%"
    .AddWidthForField "Points", "10%"
    .AddWidthForField "Tournament", "10%"
    .AddWidthForField "NextOff", "12%"
  End With
  With objReport
    .Title = "Sample Report"
    .ReportFile = App.Path & "\report.htm"
    .CssFile = App.Path & "\report.css"
    .HeaderIncFile = App.Path & "\header.inc"
    .FooterIncFile = App.Path & "\footer.inc"
    .RowsPerPage = 40
    .OpenReport
    .PrintTableFromRecordset RS, objTableConfig
    .CloseReport
    blnReportCreated = True
    wbReport.Navigate .ReportFile
  End With
  RS.Close
  lngHWnd = 0
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  Const REGPATH = "HKLM\SOFTWARE\Microsoft\Internet Explorer\AdvancedOptions\PRINT\BACKGROUND\"
  
  strPrintBackKey = "HKCU\" & _
        ReadKey(REGPATH & "RegPath") & "\" & _
        ReadKey(REGPATH & "ValueName")
  strPrintBackChecked = ReadKey(REGPATH & "CheckedValue")
  strPrintBackUnchecked = ReadKey(REGPATH & "UncheckedValue")
  chkPrintBackground.Value = IIf(ReadKey(strPrintBackKey) = strPrintBackChecked, Checked, Unchecked)
End Sub

Private Sub Timer1_Timer()
  ' Let's go to change Internet Explorer Preview Window
  If lngHWnd = 0 Then
    lngHWnd = FindWindow(vbNullString, "Print Preview")
    ' This works fine with english version of IE only... If you
    ' suspect that users can use other language version of IE,
    ' you must place additional code here:
    If lngHWnd = 0 Then
      ' Check for preview window generated by polish version of IE
      lngHWnd = FindWindow(vbNullString, "Podgl¹d wydruku")
    End If
    ' If you don't need this, simply comment above lines
  Else
    Timer1.Enabled = False
    SetWindowText lngHWnd, "HTML Reports Sample"
    ShowWindow lngHWnd, SW_MAXIMIZE
  End If
End Sub

Private Sub wbReport_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
  If Progress = 0 And blnReportCreated Then
    wbReport.ExecWB OLECMDID_PRINTPREVIEW, 0, Null, Null
  End If
End Sub

Private Function ReadKey(ByVal Key As String) As String
  ' Reads value from registry
  Dim wsh As Object
  
  On Error Resume Next
  Set wsh = CreateObject("WScript.Shell")
  ReadKey = wsh.RegRead(Key)
End Function

Private Sub WriteKey(ByVal Key As String, ByVal Value As String)
  ' Writes value to registry
  Dim wsh As Object
  
  On Error Resume Next
  Set wsh = CreateObject("WScript.Shell")
  wsh.RegWrite Key, Value
End Sub

