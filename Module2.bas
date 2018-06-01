Attribute VB_Name = "Module2"
Public con As New ADODB.connection
Public rs As New ADODB.Recordset
Public path As String

Public Sub connection()
  path = App.path & "\student.mdb"
  con.Open ("Provider=Microsoft.ACE.OLEDB.12.0;DataSource=" & path & ";Persist Security Info=False")
End Sub

