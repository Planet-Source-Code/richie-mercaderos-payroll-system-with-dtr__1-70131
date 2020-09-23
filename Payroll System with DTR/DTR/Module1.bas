Attribute VB_Name = "Module1"
Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub

