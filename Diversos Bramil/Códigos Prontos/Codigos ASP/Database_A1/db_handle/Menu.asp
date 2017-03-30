<!-- #include file="menu.cls" -->
<%
'--------------------------------------------------------------------------------------------
' Modulo Base Pagina SQL
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso del file menu.cls
'
' DSN			testo - nome connessione DSN
' SubCommand	testo - nome sotto comando
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

Dim objRS, objConn
' variabili connessione
Set objRs=Server.CreateObject("ADODB.Recordset")
Set objConn=Server.CreateObject("ADODB.Connection")

Dim annulla
annulla = False

Dim MyName, MySQL, MyIDValue, MyDSN, MySubCommand, MyMainRitorno

' nome
MyName = SubCommand

' prelevamento valori campi modulo per DSN e SubCommand se esistono
MyDSN = cstr(request.form("DSN"))

' prelevamento valori campi modulo per DSN e SubCommand se esistono
MySubCommand = cstr(request.form("SubCommand"))

' prelevamento valori campi modulo per DSN e SubCommand se esistono
MyMainRitorno = cstr(request.form("MainRitorno"))

' controlli sullo switch per i campi d'ingresso obbligatori
if MySubCommand="" then
	MySubCommand=SubCommand
	if MySubCommand="" then
		annulla=True
		MyErr = "Error: missing Command"
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
	MySQL = "SELECT " & MainTable & ".* FROM " & MainTable & " WHERE (((" & MainTable & ".Nome)='" & MySubCommand & "'));"

	objConn.open MyDSN,"",""
	' apertura dati db
	set objRS = objConn.Execute(MySQL,intRecordAffected)

	' controllo esistenza dati
	if intRecordAffected <0 Then
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

	MySQL = "SELECT " & SubTable & ".* FROM " & SubTable & " WHERE (((" & SubTable & ".CommandID)=" & MyIDValue & ")) ORDER BY " & SubTable & ".TableNome;"
end if
%>
<html>

<head>
<title><%Response.write(MyName)%></title>

<meta name="Microsoft Border" content="none"></head>

<body style="font-family: Arial">
<div align="center">

<p align="center"><img src="logo.gif" WIDTH="81" HEIGHT="78"><font face="Courier" color="#000080"><big><big><strong>
dB Administration</strong></big></big></font></p>

<p align="center"><font color="#000080" face="Arial Narrow"><%if MyMainRitorno<>"" then%>[ <a href="<%Response.write(MyMainRitorno)%>">Up</a> ]<%end if%></font></p>

<p align="center"><big><big><font face="Courier" color="#000080"><strong><%Response.write("Menu " & MyName)%></strong></font></big></big><%' Interruzzione per input dati errato
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
		objRS.MoveFirst%><br>
</p>
<div align="center"><center>

<table border="0" cellpadding="0">
  <tr>
    <td></td>
  </tr>
<%Do while Not objRS.EOF%>
  <tr>
    <td align="center"><form method="POST" action="sql_table.asp" style="font-family: monospace; margin: 0px; padding: 0px">
      <input type="hidden" name="DSN" value="<%Response.write(objRS.Fields("TableDSN").Value)%>"><input type="hidden" name="Table" value="<%Response.write(objRS.Fields("TableName").Value)%>"><input type="hidden" name="SQL" value="<%Response.write(objRS.Fields("TableSQL").Value)%>"><input type="hidden" name="IDName" value="<%Response.write(objRS.Fields("TableIDName").Value)%>"><input type="hidden" name="Titolo" value="<%Response.write(objRS.Fields("TableTitolo").Value)%>"><input type="hidden" name="Comandi" value="<%Response.write(objRS.Fields("TableComandi").Value)%>"><input type="hidden" name="Ritorno" value="<%Response.write(Request.ServerVariables("PATH_INFO"))%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MySubCommand)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><p><input type="submit" value="<%Response.write(objRS.Fields("TableNome").Value)%>" name="B1" style="font-family: monospace; margin: 0px; padding: 0px"></p>
    </form>
    </td>
  </tr>
<%objRS.MoveNext
Loop
objRS.Close
end if
end if%>
  <tr>
    <td><p align="center"><font color="#000080" face="Arial Narrow"><strong><%if MyMainRitorno="" then%> <%Response.write("<br>Chiudere la finestra")
end if%></strong></font></td>
  </tr>
</table>
</center></div></div>

<p><%
Set annulla=Nothing
Set MyErr=Nothing

Set MyName=Nothing
Set MySQL=Nothing
Set MyIDValue=Nothing
Set MyDSN=Nothing
Set MySubCommand=Nothing
Set MyMainRitorno=Nothing

Set objRS=Nothing
Set objConn=Nothing

Session.Abandon
%> </p>
</body>
</html>
