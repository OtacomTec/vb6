<!-- #include file="main.cls" -->
<%
'--------------------------------------------------------------------------------------------
' Modulo Pagina Principale SQL
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso del file main.cls
'
' DSN			testo - nome connessione DSN
' Command		testo - nome tabella command
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

Dim objRS, objConn
' variabili connessione
Set objRs=Server.CreateObject("ADODB.Recordset")
Set objConn=Server.CreateObject("ADODB.Connection")

Dim annulla
annulla = False

Dim MyName, MySQL, MyIDValue, MyDSN, MyCommand

' nome
MyName = SubCommand

' prelevamento valori campi modulo per DSN e SubCommand se esistono
MyDSN = cstr(request.form("DSN"))

' controlli sullo switch per i campi d'ingresso obbligatori
if MyCommand="" then
	MyCommand=Command
	if MyCommand="" then
		annulla=True
		MyErr = "Error: missing command"
	end if
end if
if MyDSN="" then
	MyDSN = DSN
	if MyDSN="" then
		annulla=True
		MyErr = "Error: missing DSN"
	end if
end if

if Not annulla then
	annulla = TRUE
	MySQL = "SELECT " & MyCommand & ".* FROM " & MyCommand & " ORDER BY " & MyCommand & ".Nome;"

	objConn.open MyDSN,"",""
	' apertura dati db

	' apertura dati db
	objRS.Open MySQL, objConn, 2, 1

	' controllo esistenza dati
	if Not objRS.EOF then
		objRS.MoveFirst
		MyIDValue = objRS.Fields("CommandID").Value
		MyName = objRS.Fields("Nome").Value
		If Not IsNull(MyIDValue) then
			annulla=FALSE
		end if
		objRS.Close
	end if

	Set objRS=Nothing
	Set objConn=Nothing
end if
%>
<html>

<head>
<title><%Response.write(MyName)%></title>

<meta name="Microsoft Border" content="none"></head>

<body>

<p align="center"><img src="logo.gif" WIDTH="81" HEIGHT="78">&nbsp; <big><big><font face="Courier" color="#000080"><strong>dB Administration</strong></font></big></big></p>

<p align="center"><font color="#000080" face="Arial Narrow">[ <a href="<%Response.write(MainBack)%>">Up</a> ]</font></p>

<p><%
' Interruzzione per input dati errato
if annulla then
	Response.Write("<CENTER><H4>ATTENTION</H4></CENTER>")
	Response.Write("<CENTER>" & MyErr & "</CENTER><br>")
else
	Set objRs=Server.CreateObject("ADODB.Recordset")
	Set objConn=Server.CreateObject("ADODB.Connection")

	' apertura connessione
	objConn.open MyDSN,"",""

	' apertura dati db
	objRS.Open MySQL, objConn, 2, 1

	' controllo esistenza dati
	if Not objRS.EOF then
		objRS.MoveFirst%> </p>
<div align="center"><center>

<table border="0" cellpadding="0">
<%do while Not objRS.EOF%>
  <tr>
    <td align="center"><form method="POST" action="Menu.asp" style="font-family: monospace; margin: 0px; padding: 0px">
      <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="SubCommand" value="<%Response.write(objRS.Fields("Nome").Value)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(Request.ServerVariables("PATH_INFO"))%>"><p><input type="submit" value="<%Response.write(objRS.Fields("Nome").Value)%>" name="B1" style="font-family: monospace"></p>
    </form>
    </td>
  </tr>
<%objRS.MoveNext
Loop
objRS.Close
end if
end if%>
</table>
</center></div>

<p>&nbsp; </p>

<p><%
Set annulla=Nothing
Set MyErr=Nothing

Set MyName=Nothing
Set MySQL=Nothing
Set MyIDValue=Nothing
Set MyDSN=Nothing
Set MyCommand=Nothing

Set objRS=Nothing
Set objConn=Nothing

Session.Abandon
%> </p>
</body>
</html>
