<%
' Modulo classe Esecuzione SQL Update, Insert e Delete
' 
' Parametri d'ingresso
' clsRunSQL.DSN - DSN di sistema
' clsRunSQL.SQL - istruzione SQL
'
' Parametri d'uscita
' clsRunSQL.Result - risultato SQL
'

CLASS clsRunSQL

Private MyDSN
Private MySQLCmd

Private function Esegui_SQL
On Error Resume Next

Dim objConn, objRS
Dim Stringa, Fase

Fase=0

Stringa = "dsn=" & MyDSN & ";database=" & MyDSN

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Stringa
'Set Session("My_conn") = objConn

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open MySQLCmd, objConn, 3, 3

Fase=1
Stringa=""

objRS.Close
'controllo errori

if Err.Description <> "" then
	if Fase=0 then
		Stringa = Err.Description
	end if
end if

Set objRS=Nothing
Set objConn=Nothing

if Stringa <> "" then
	Stringa = "Error: " & Stringa
else
	Stringa = "Status Operation: done"
end if

Esegui_SQL = Stringa

Set Stringa = Nothing
Set Fase = Nothing
End function


' Impostazione parametro d'ingresso DSN
Public Property Let DSN(NewDSN)
MyDSN = NewDSN
End Property

' Impostazione parametro d'ingresso SQL
Public Property Let SQL(NewSQLCmd)
MySQLCmd = NewSQLCmd
End Property

' Impostazione parametri d'uscita
Public Property Get Result
Result = Esegui_SQL
End Property

END CLASS
%>
