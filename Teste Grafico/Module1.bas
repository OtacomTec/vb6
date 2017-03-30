Attribute VB_Name = "Module1"
Option Explicit

Public cn As New ADODB.Connection

Sub Main()
    
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Teste Grafico\NWIND.MDB;Persist Security Info=False"
    cn.CursorLocation = adUseClient
    cn.Open
    
    Form1.Show
    
End Sub

Public Function sp_Transacoes(ByVal pstrOpcao As String) As ADODB.Recordset
    
    Dim strSQL      As String

    Select Case pstrOpcao
    
        Case "PES"
            strSQL = "SELECT TOP 10 * FROM ORDERS"
        
    End Select
    
    
    Set sp_Transacoes = cn.Execute(strSQL)
    
End Function

