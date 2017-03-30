<%@ Language=VBScript %>
<!-- #include file="sql_field.cls" -->
<%
'--------------------------------------------------------------------------------------------
' Modulo Eliminazione Pagina automatica accesso ai dati di una tabella di un database
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso
'
' DSN			testo - nome connessione DSN
' Table		testo - nome tabella - il primo campo deve essere l'identificativo del record
' Titolo		testo - titolo pagina - se omesso verra' utilizzata la dicitura Dati db
' Ritorno		testo - pagina di ritorno - se omesso verra' indicato Chiudere la finestra
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

' classe caratteristiche Field
Dim MyObjField
Set MyObjField = New clsObjField
' modalita' aggiunta nuovo
MyObjField.SetMode = True

' variabili connessione
Dim objRS, objConn
Set objRs=Server.CreateObject("ADODB.Recordset")
Set objConn=Server.CreateObject("ADODB.Connection")

' flag annullamento operazione
Dim annulla
annulla=False

' codice errore
Dim MyErr

' variabili di ingresso
Dim MyDSN, MyTable, MyRetSQL, MySQL, MySQLCmd, MyIDName, MyIDValue, MyTitolo, MyRitorno, MyRetTitolo, MyRetRitorno, MyRetComandi, MyMainDSN, MyMainSubCommand, MyMainRitorno

' lettura valori DSN e tabella e controllo degli errori
' 1 - nome DSN
MyDSN = cstr(request.form("DSN"))

' 2 - nome tabella
MyTable = cstr(request.form("Table"))

' 3 - ret sql
MyRetSQL = cstr(request.form("RetSQL"))

' 4 - ID Name
MyIDName = cstr(request.form("IDName"))

' 5 - ID Value
MyIDValue = cstr(request.form("IDValue"))

' 6 - titolo
MyTitolo = cstr(request.form("Titolo"))

' 7 - ritorno
MyRitorno = cstr(request.form("Ritorno"))

' 8 - ret titolo
MyRetTitolo = cstr(request.form("RetTitolo"))

' 9 - ret ritorno
MyRetRitorno = cstr(request.form("RetRitorno"))

' 10 - ret comandi
MyRetComandi = cstr(request.form("RetComandi"))

' 11 - main DSN
MyMainDSN = cstr(request.form("MainDSN"))

' 12 - main SubCommand
MyMainSubCommand = cstr(request.form("MainSubCommand"))

' 13 - main ritorno
MyMainRitorno = cstr(request.form("MainRitorno"))

' controlli sui campi di ingresso obbligatori
if MyDSN="" then
	annulla=True
	MyErr = "Error: missing DSN"
elseif MyTable="" or MyIDValue= "" or MyIDName="" then
	annulla=True
	MyErr = "Error: missing Table Name"
end if

' controllo sul titolo
if MyTitolo="" then
	MyTitolo = "Delete dB Data"
else
	MyTitolo = "Delete " & MyTitolo
end if

MySQL = "SELECT " & MyTable & ".* FROM " & MyTable & " WHERE (((" & MyTable & "." & MyIDName & ")=" & MyIDValue & "));"
MySQLCmd = "DELETE " & MyTable & ".*, " & MyTable & "." & MyIDName & " FROM " & MyTable & " WHERE (((" & MyTable & "." & MyIDName & ")=" & MyIDValue & "));"
%>
<html>

<head>
<title><%Response.write(MyTitolo)%></title>
<meta name="Microsoft Border" content="none">
</head>

<body style="font-family: Arial">

<p align="center"><img src="logo.gif" width="81" height="78"><font face="Courier"
color="#000080"><big><big><strong> dB Administration - <%Response.write(MyTitolo)%></strong></big></big></font><br>
</p>

<form method="POST" action="sql_cmd.asp" target="_blank">
  <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden"
  name="SQL" value="<%Response.Write(MySQLCmd)%>"><input type="hidden" name="Titolo"
  value="<%Response.Write(MyTitolo)%>"><div align="center"><center><p><input type="submit"
  value="Confirm" name="Conferma" style="font-family: Arial; font-size: 8pt"></p>
  </center></div>
</form>

<p align="center"><%
' Interruzzione per input dati errato
if annulla then
	Response.Write("<CENTER><H4>ATTENTION</H4></CENTER>")
	Response.Write("<CENTER>" & MyErr & "</CENTER><br>")
else
	objConn.open MyDSN,"",""
	objRS.Open MySQL, objConn, 2, 1

	' segue disegno tabella
%></p>
<div align="center"><center>

<table BORDER="<%=objRS.Fields.Count%>" COLS="2">
  <tr>
<%	For each objField in objRS.Fields
		If objField.Name<>MyIDName then

			' imposto i valori per la classe
			MyObjField.SetField = objField

			If MyObjField.Updatable then%>
    <td align="right"><font face="Arial Narrow"><%=objField.Name%></font></td>
    <td align="left"><font face="Arial Narrow"><%=MyObjField.NewParser%></font></td>
  </tr>
<%			End if
		End if %>
<%	Next %>
</table>
</center></div><div align="center">

<p align="center"><%' chiusura
	objRS.close
end if
%></p>

<p align="center"><%if MyRitorno<>"" then%> </p>

<form method="POST" action="<%Response.Write(MyRitorno)%>">
  <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden"
  name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="SQL"
  value="<%Response.write(MyRetSQL)%>"><input type="hidden" name="IDName"
  value="<%Response.write(MyIDName)%>"><input type="hidden" name="Titolo"
  value="<%Response.write(MyRetTitolo)%>"><input type="hidden" name="Ritorno"
  value="<%Response.write(MyRetRitorno)%>"><input type="hidden" name="Comandi"
  value="<%Response.write(MyRetComandi)%>"><input type="hidden" name="MainDSN"
  value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno"
  value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand"
  value="<%Response.write(MyMainSubCommand)%>"><div align="center"><center><p><input
  type="submit" value="Back" name="Ritorna" style="font-size: 8pt"> </p>
  </center></div>
</form>
<%else
	Response.Write("Close the window")
end if
%>

<p align="center"><%
Set annulla=Nothing
Set MyErr=Nothing

Set MyDSN=Nothing
Set MyTable=Nothing
Set MySQL=Nothing
Set MySQLCmd=Nothing
Set MyRetSQL=Nothing
Set MyIDName=Nothing
Set MyIDValue=Nothing
Set MyTitolo=Nothing
Set MyRitorno=Nothing
Set MyRetTitolo=Nothing
Set MyRetRitorno=Nothing
Set MyRetComandi=Nothing
Set MyMainDSN=Nothing
Set MyMainSubCommand=Nothing
Set MyMainRitorno=Nothing

Set MyObjField=Nothing

Set objRS=Nothing
Set objConn=Nothing

Session.Abandon
%> </p>

<p align="center">&nbsp;</p>
</div>

<p>&nbsp; </p>
</body>
</html>
