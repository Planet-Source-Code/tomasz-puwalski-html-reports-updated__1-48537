VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------
' HTML Reports - TableConfig class
' Created by Tomasz Puwalski (pvl@cps.pl)
' Version 1.0 (09/16/2003)
'---------------------------------------------
Option Explicit

Private Headers As New Dictionary
Private HAligns As New Dictionary
Private Widths As New Dictionary

Enum HAlign
  [Left Justify] = 0
  [Right Justify] = 1
  [Center] = 2
End Enum


' Properties
Public Border As Integer
Public CellPadding As Integer
Public CellSpacing As Integer
Public Class As String
Public Height As String
Public Width As String


' Methods
Public Sub AddHeaderForField(ByVal FieldName As String, ByVal Header As String)
  If Headers.Exists(FieldName) Then
    Headers(FieldName) = Header
  Else
    Headers.Add FieldName, Header
  End If
End Sub

Public Function GetHeaderForField(ByVal FieldName As String) As String
  If Headers.Exists(FieldName) Then
    GetHeaderForField = Headers(FieldName)
  Else
    GetHeaderForField = FieldName
  End If
End Function

Public Sub AddAlignForField(ByVal FieldName As String, ByVal Alignment As HAlign)
  If HAligns.Exists(FieldName) Then
    HAligns(FieldName) = Alignment
  Else
    HAligns.Add FieldName, Alignment
  End If
End Sub

Public Function GetAlignForField(ByVal FieldName As String) As String
  If HAligns.Exists(FieldName) Then
    Select Case HAligns(FieldName)
    Case [Left Justify]
      GetAlignForField = "left"
    Case [Right Justify]
      GetAlignForField = "right"
    Case [Center]
      GetAlignForField = "center"
    End Select
  Else
    GetAlignForField = vbNullString
  End If
End Function

Public Sub AddWidthForField(ByVal FieldName As String, ByVal WidthValue As String)
  If Widths.Exists(FieldName) Then
    Widths(FieldName) = WidthValue
  Else
    Widths.Add FieldName, WidthValue
  End If
End Sub

Public Function GetWidthForField(ByVal FieldName As String) As String
  If Widths.Exists(FieldName) Then
    GetWidthForField = Widths(FieldName)
  Else
    GetWidthForField = vbNullString
  End If
End Function

