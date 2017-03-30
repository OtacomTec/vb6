<%@ Language=VBScript %>
<!-- #include file="sql_cmd.cls" -->
<%
'--------------------------------------------------------------------------------------------
' Modulo Esecuzione istruzione SQL
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso
'
' DSN			testo - nome connessione DSN
' SQL			testo - espressione SQL
' Titolo		testo - titolo pagina - se omesso verra' utilizzata la dicitura Conferma
'						 Operazione Dati db
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

' classe esecuzione SQL
Dim MyRunSQL
Set MyRunSQL = New clsRunSQL

' flag annullamento operazione
Dim annulla
annulla=False

' codice errore
Dim MyErr

' variabili di ingresso
Dim MyDSN, MySQL, MyTitolo

' lettura valori DSN e tabella e controllo degli errori
' 1 - nome DSN
MyDSN = cstr(request.form("DSN"))

' 2 - istruzione SQL
MySQL = cstr(request.form("SQL"))

' 3 - titolo
MyTitolo = cstr(request.form("Titolo"))

' controlli sui campi di ingresso obbligatori
if MyDSN="" then
	annulla=True
	MyErr = "Error: missing DSN"
end if

' controllo sulla query specifica
if MySQL="" then
	annulla=True
	MyErr = "Error: missing Command"
end if

' controllo sul titolo
if MyTitolo="" then
	MyTitolo = "Confirm dB Data Operation"
else
	MyTitolo = "Confirm " & MyTitolo
end if
%>
<html>

<head>
<title><%Response.write(MyTitolo)%></title>

<meta name="Microsoft Border" content="none"></head>

<body style="font-family: Arial">

<p align="center"><img src="logo.gif" WIDTH="81" HEIGHT="78"><font face="Courier" color="#000080"><big><big><strong>&nbsp;
<%Response.write(MyTitolo)%></strong></big></big></font><br>
</p>

<p align="center"><%
' Interruzzione per input dati errato
if annulla then
	Response.Write("<CENTER><H4>ATTENTION</H4></CENTER>")
	Response.Write("<CENTER>" & MyErr & "</CENTER><br>")
else

MyRunSQL.DSN = MyDSN
MyRunSQL.SQL = MySQL
MyErr = MyRunSQL.Result
Response.Write(MyErr)
end if

Set annulla=Nothing
Set MyErr=Nothing

Set MyDSN=Nothing
Set MySQL=Nothing
Set MyTitolo=Nothing
Set MyRunSQL = Nothing

Session.Abandon
%></p>

<p align="center"><font color="#000080" face="Arial Narrow"><big>Close the window</big> </font></p>
</body>
</html>
