Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Enum RecOp
    rEdit = 0
    rAddNew = 1
End Enum

Public EditOp As RecOp
Public ButtonPress As VbMsgBoxResult

Public FindFieldIdx As Integer
Public FindValue As String

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    If Len(lzFileName) = 0 Then Exit Function
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function GetFieldSize(ByVal fName As String) As Byte
    'Return fieldsize
    Select Case fName
        Case "Boolean"
            GetFieldSize = 1
        Case "Byte"
            GetFieldSize = 1
        Case "Integer"
            GetFieldSize = 2
        Case "Long"
            GetFieldSize = 4
        Case "Currency"
            GetFieldSize = 8
        Case "Single"
            GetFieldSize = 4
        Case "Double"
            GetFieldSize = 8
        Case "Date/Time"
            GetFieldSize = 8
        Case "Text"
            GetFieldSize = 50
        Case "Memo"
            GetFieldSize = 0
    End Select
End Function

Function GetFieldType(ByVal fName As String) As Integer
    'Return FieldType
    Select Case fName
        Case "Boolean"
            GetFieldType = dbBoolean
        Case "Byte"
            GetFieldType = dbByte
        Case "Integer"
            GetFieldType = dbInteger
        Case "Long"
            GetFieldType = dbLong
        Case "Currency"
            GetFieldType = dbCurrency
        Case "Single"
            GetFieldType = dbSingle
        Case "Double"
            GetFieldType = dbDouble
        Case "Date/Time"
            GetFieldType = dbDate
        Case "Text"
            GetFieldType = dbText
        Case "Memo"
            GetFieldType = dbMemo
    End Select
End Function

Function GetFieldTypeStr(ByVal fName As String) As String
    'Return FieldType string
    Select Case fName
        Case dbBoolean
            GetFieldTypeStr = "Boolean"
        Case dbByte
            GetFieldTypeStr = "Byte"
        Case dbInteger
            GetFieldTypeStr = "Integer"
        Case dbLong
            GetFieldTypeStr = "Long"
        Case dbCurrency
            GetFieldTypeStr = "Currency"
        Case dbSingle
            GetFieldTypeStr = "Single"
        Case dbDouble
            GetFieldTypeStr = "Double"
        Case dbDate
            GetFieldTypeStr = "Date"
        Case dbText
            GetFieldTypeStr = "Text"
        Case dbMemo
            GetFieldTypeStr = "Memo"
    End Select
End Function

Public Function RunApp(iHwnd As Long, OpenOp As String, FileName As String) As Long
    RunApp = ShellExecute(iHwnd, OpenOp, FileName, "", "", 1)
End Function

