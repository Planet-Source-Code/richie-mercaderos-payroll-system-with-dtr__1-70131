Attribute VB_Name = "Module1"
Option Explicit

Global cn               As New ADODB.Connection
Global rs               As New ADODB.Recordset
Global rst               As New ADODB.Recordset

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

Public Function LoadListview1(ByRef sRecord As ADODB.Recordset, ByVal sListview As ListView)
On Error Resume Next
Dim x As ListItem

sRecord.Requery

Do Until sRecord.EOF
    Set x = sListview.ListItems.Add(, , "", 1, 1)
        x.SubItems(1) = sRecord.Fields("No")
        x.SubItems(2) = sRecord.Fields("Name")
    
    sRecord.MoveNext
Loop
End Function

