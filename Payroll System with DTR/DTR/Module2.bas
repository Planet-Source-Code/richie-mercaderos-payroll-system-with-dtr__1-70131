Attribute VB_Name = "Module2"
Option Explicit

Global cn               As New ADODB.Connection
'----------------------------------------------
Global rs               As New ADODB.Recordset
Global time_rec         As New ADODB.Recordset
'----------------------------------------------

Public Type PData

    fname As String
    pmonth As String

End Type

Public pview As PData

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
    MsgBox "Please type employee number.", vbExclamation, "Data Requirements"
    sText.SetFocus
End If
End Function
