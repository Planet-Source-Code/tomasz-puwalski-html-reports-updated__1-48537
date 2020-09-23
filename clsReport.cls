VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------
' HTML Reports - Report class
' Created by Tomasz Puwalski (pvl@cps.pl)
' Version 1.0 (09/16/2003)
'---------------------------------------------
' Properties:
'   CssFile - Cascading Style Sheets file name
'   DeleteReportOnExit -
'   FooterIncFile - footer file name (include)
'   HeaderIncFile - header file name (include)
'   ReportFile - report file name
'   RowsPerPage - number of table rows per page
'   Title - report title
'---------------------------------------------
' Methods:
'   CloseReport - ends report generation
'   OpenReport - begins report generation
'   PrintTableFromRecordset - prints data
'---------------------------------------------
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private blnReportCreated As Boolean
Private intFileNum As Integer
Private strHeaderIncFile As String
Private strHeader As String
Private strFooterIncFile As String
Private strFooter As String

' Public properties (with full read/write rights)
Public ReportFile As String
Public CssFile As String
Public DeleteReportOnExit As Boolean
Public RowsPerPage As Integer
Public Title As String

' Other properties
Public Property Get HeaderIncFile() As String
  HeaderIncFile = strHeaderIncFile
End Property

Public Property Let HeaderIncFile(NewValue As String)
  If FileExists(NewValue) Then
    strHeader = GetFileContent(NewValue)
    strHeaderIncFile = NewValue
  End If
End Property

Public Property Get FooterIncFile() As String
  FooterIncFile = strFooterIncFile
End Property

Public Property Let FooterIncFile(NewValue As String)
  If FileExists(NewValue) Then
    strFooter = GetFileContent(NewValue)
    strFooterIncFile = NewValue
  End If
End Property


' Methods
Public Sub OpenReport()
  If ReportFile <> vbNullString And intFileNum = 0 Then
    If FileExists(ReportFile) Then
      Kill ReportFile
    End If
    intFileNum = FreeFile
    Open ReportFile For Append As #intFileNum
    Print #intFileNum, "<html><head>"
    If Title <> vbNullString Then
      Print #intFileNum, "<title>" & Title & "</title>"
    End If
    If CssFile <> vbNullString Then
      Print #intFileNum, "<link rel=""stylesheet"" href=""file:///" & CssFile & """ type=""text/css"">"
    End If
    Print #intFileNum, "<script type=""Text/JavaScript"" language=""JavaScript"">"
    Print #intFileNum, "<!--"
    Print #intFileNum, "pageno = 1;"
    Print #intFileNum, "//-->"
    Print #intFileNum, "</script>"
    Print #intFileNum, "</head>"
    Print #intFileNum, "<body><div>"
    If strHeader <> vbNullString Then
      Print #intFileNum, strHeader
    End If
  End If
End Sub

Public Sub CloseReport()
  If intFileNum <> 0 Then
    If strFooter <> vbNullString Then
      Print #intFileNum, strFooter
    End If
    Print #intFileNum, "</div></body></html>"
    Close #intFileNum
  End If
End Sub

Public Sub PrintTableFromRecordset(Source As Recordset, Config As clsTableConfig, Optional ByVal WithHeader As Boolean = True)
  Dim fld As Object
  Dim blnEven As Boolean
  Dim strAlignment As String
  Dim strWidth As String
  Dim intRowNo As Integer
  
  If intFileNum <> 0 Then
    PrintTableHeader Source, Config, WithHeader
    Do While Not Source.EOF
      intRowNo = intRowNo + 1
      Print #intFileNum, "<tr class=""" & IIf(blnEven, "DataRowEven", "DataRowOdd") & """>"
      For Each fld In Source.Fields
        strAlignment = Config.GetAlignForField(fld.Name)
        strWidth = Config.GetWidthForField(fld.Name)
        Print #intFileNum, "<td class=""DataCell""" & _
            IIf(strAlignment = vbNullString, "", " align=""" & strAlignment & """") & _
            IIf(strWidth = vbNullString, "", " width=""" & strWidth & """") & _
            ">&nbsp;" & CStr(fld.Value) & "&nbsp;</td>"
      Next
      Print #intFileNum, "</tr>"
      Source.MoveNext
      If RowsPerPage > 0 Then
        If intRowNo = RowsPerPage And Not Source.EOF Then
          Print #intFileNum, "</table></center>"
          If strFooter <> vbNullString Then
            Print #intFileNum, strFooter
          End If
          Print #intFileNum, "</div>"
          Print #intFileNum, "<div style=""page-break-before:always"">" ' ;visibility:hidden
          If strHeader <> vbNullString Then
            Print #intFileNum, strHeader
          End If
          PrintTableHeader Source, Config, WithHeader
          intRowNo = 0
        End If
      End If
      blnEven = Not blnEven
    Loop
    Print #intFileNum, "</table></center>"
  End If
End Sub


' Internal procedures and functions
Private Sub Class_Initialize()
  Dim strTempFile As String
  Dim strTempPath As String
  Dim intN As Integer
  
  Title = "Sample Report"
  DeleteReportOnExit = True
  ' Let's create file name
  strTempPath = String(100, Chr(0))
  Call GetTempPath(100, strTempPath)
  strTempPath = Left(strTempPath, InStr(strTempPath, Chr(0)) - 1) & "HtmlReport\"
  ReportFile = strTempPath
  Randomize
  For intN = 1 To 8
    ReportFile = ReportFile & Chr(Asc("A") + Rnd() * 25)
  Next
  ReportFile = ReportFile & ".htm"
End Sub

Private Sub Class_Terminate()
  If blnReportCreated Then
    On Error Resume Next
    Kill ReportFile
  End If
End Sub

Private Function FileExists(ByVal FileName As String)
  Dim a As VbFileAttribute
  
  On Error Resume Next
  a = GetAttr(FileName)
  If Err.Number <> 0 Then
    Exit Function
  End If
  FileExists = Not (a And vbDirectory)
End Function

Private Function GetFileContent(ByVal FileName As String) As String
  Dim intFileNum As Integer
  Dim strLine As String
    
  intFileNum = FreeFile
  Open FileName For Input As #intFileNum
  Do While Not EOF(intFileNum)
    Line Input #intFileNum, strLine
    If GetFileContent <> vbNullString Then
      GetFileContent = GetFileContent & vbNewLine
    End If
    GetFileContent = GetFileContent & strLine
  Loop
  Close #intFileNum
End Function

Private Sub PrintTableHeader(Source As Recordset, Config As clsTableConfig, ByVal WithHeader As Boolean)
  Dim fld As Object
  
  Print #intFileNum, "<center><table border=" & Config.Border
  If Config.Height <> vbNullString Then
    Print #intFileNum, "height=" & Config.Height
  End If
  If Config.Width <> vbNullString Then
    Print #intFileNum, "width=" & Config.Width
  End If
  Print #intFileNum, " cellpadding=" & Config.CellPadding & " cellspacing=" & Config.CellSpacing & " class=""" & Config.Class & """>"
  If WithHeader Then
    Print #intFileNum, "<tr class=""DataHeader"">"
    For Each fld In Source.Fields
      Print #intFileNum, "<td class=""DataHeader"">&nbsp;" & Config.GetHeaderForField(fld.Name) & "&nbsp;</td>"
    Next
    Print #intFileNum, "</tr>"
  End If
End Sub

