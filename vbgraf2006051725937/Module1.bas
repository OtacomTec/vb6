Attribute VB_Name = "Module1"
Global com As ADODB.Connection


Sub main()
conectar
Form1.Show
End Sub

Public Sub conectar()
Set com = New ADODB.Connection
com.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\biblio.mdb" & ";Persist Security Info=False;"
com.ConnectionTimeout = 1000
com.CursorLocation = adUseClient
com.Open
End Sub
