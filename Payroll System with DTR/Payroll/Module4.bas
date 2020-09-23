Attribute VB_Name = "Module4"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public sTimeSet As Integer

Public Function sLoginUser(ByRef sRecordset As ADODB.Recordset, ByRef sField As String, ByVal sText As String) As Boolean
sRecordset.Requery
sRecordset.Find sField & "='" & sText & "'"

If sRecordset.EOF Then
    sLoginUser = False
Else
    sLoginUser = True
End If
End Function
