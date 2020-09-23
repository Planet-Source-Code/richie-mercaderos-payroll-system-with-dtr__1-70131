Attribute VB_Name = "Module1"
Option Explicit

Public Sub set_conn_getData(ByRef sConnection As ADODB.Connection, ByVal sDataLocation As String, ByVal sHavePassword As Boolean, ByVal sPassword As String)
If sHavePassword = True Then
    sConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataLocation & ";Persist Security Info=False;Jet OLEDB:Database Password=" & sPassword
Else
    sConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataLocation & ";Persist Security Info=False"
End If
End Sub

Public Sub set_rec_getData(ByRef sRecordset As ADODB.Recordset, ByRef sConnection As ADODB.Connection, ByVal sSQL As String)
With sRecordset
    .CursorLocation = adUseClient
    .Open sSQL, sConnection, adOpenKeyset, adLockOptimistic
End With
End Sub

Public Function isempty(ByVal sText As Variant) As Boolean
If sText.Text = "" Then
    isempty = True
    MsgBox "The field must be fill.", vbExclamation, "Data Requirements"
    sText.SetFocus
End If
End Function

Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub

Public Function if_exist(ByVal sTable As String, ByVal sField As String, ByRef sEntryField As Variant) As Boolean
Dim rs As New ADODB.Recordset
if_exist = False
Call set_rec_getData(rs, cn, "Select * From " & sTable & " Where " & sField & " ='" & sEntryField.Text & "'")
If rs.RecordCount > 0 Then
    MsgBox "The adding of new entry cannot be done because '" & sEntryField.Text & "' is already" & vbCrLf & "exist in the record.Please check and change it." & vbCrLf & vbCrLf & "Note: Duplication of entries is not allowed in this form.", vbExclamation, "Unable to Add"
    sEntryField.SetFocus
    if_exist = True
End If
Set rs = Nothing
End Function

