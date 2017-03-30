Attribute VB_Name = "ODBC"
Public cn As ADODB.Connection
Public Function DBConnect() As Boolean

On Error GoTo OpenErr

Dim MSDatabase

Set cn = New ADODB.Connection

MSDatabase = App.Path & "\" & "Database.mdb"
    cn.CursorLocation = adUseClient
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open MSDatabase, Admin
    DBConnect = True
Exit Function

OpenErr:

    MsgBox "Error Opening " & MSDatabase & vbNewLine & Err.Description, vbCritical, "Open Database Error"
    DBConnect = False


End Function
