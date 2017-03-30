<%@ Language=VBScript %>
<!-- #include file="sql_field.cls" -->
<%
'--------------------------------------------------------------------------------------------
' Modulo Creazione Pagina automatica accesso ai dati di una tabella di un database
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso
'
' DSN			testo - nome connessione DSN
' Table		testo - nome tabella - il primo campo deve essere l'identificativo del record
' SQL			testo - espressione SQL - se nulla viene eseguita una query di accesso
'						 completo alla tabella MyTable
' IDName		testo - nome campo chiave unica
' Titolo		testo - titolo pagina - se omesso verra' utilizzata la dicitura Dati db
' Ritorno		testo - pagina di ritorno - se omesso verra' indicato Chiudere la finestra
' Comandi		testo - 1: aggiungi - 2: modifica - 4: cancella con tutte le combinazioni
'						 possibili (da 0 a 7)
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

' spostamento pagina
Dim MyStart
Dim MyPage
Dim MyPages
Dim MyContinua

' classe caratteristiche Field
Dim MyObjField
Set MyObjField = New clsObjField
' modalita' aggiunta nuovo
MyObjField.SetMode = True
MyObjField.SetSmartRef = True

' variabili connessione
Dim objRS, objConn
Set objRS=Server.CreateObject("ADODB.Recordset")
Set objConn=Server.CreateObject("ADODB.Connection")

' flag annullamento operazione
Dim annulla
annulla=False

' codice errore
Dim MyErr

' variabili di ingresso
Dim MyDSN, MyTable, MySQL, MyIDName, MyTitolo, MyRitorno, MyModCmd, MyAggiungi, MyModifica, MyCamcella, MyAbsPath, MyCmd, MyMainDSN, MyMainSubCommand, MyMainRitorno

' lettura valori DSN e tabella e controllo degli errori
' 1 - nome DSN
MyDSN = cstr(request.form("DSN"))

' 2 - nome tabella
MyTable = cstr(request.form("Table"))

' 3 - istruzione SQL
MySQL = cstr(request.form("SQL"))

' 4 - IDName
MyIDName = cstr(request.form("IDName"))

' 5 - titolo
MyTitolo = cstr(request.form("Titolo"))

' 6 - ritorno
MyRitorno = cstr(request.form("Ritorno"))

' 7 - comandi
MyModCmd = cstr(request.form("Comandi"))

' 8 - main DSN
MyMainDSN = cstr(request.form("MainDSN"))

' 9 - main SubCommand
MyMainSubCommand = cstr(request.form("MainSubCommand"))

' 10 - main ritorno
MyMainRitorno = cstr(request.form("MainRitorno"))

' controlli sui campi di ingresso obbligatori
if MyDSN="" then
	annulla=True
	MyErr = "Error: missing DSN"
elseif MyTable="" then
	annulla=True
	MyErr = "Error: missing Table Name"
end if

' controllo sulla query specifica
if MySQL="" then
	MySQL = "SELECT " & MyTable & ".* FROM " & MyTable & ";"
end if

' controllo sul titolo
if MyTitolo="" then
	MyTitolo = "Data"
end if

' controllo sull'abilitazione per i comandi di modifica e l'esistenza del campo primario
MyAggiungi = False
MyModifica = False
MyCancella = False

SELECT CASE MyModCmd
Case "1"
	MyAggiungi=True
Case "2"
	MyModifica=True
Case "3"
	MyAggiungi=True
	MyModifica=True
Case "4"
	MyCancella=True
Case "5"
	MyAggiungi=True
	MyCancella=True
Case "6"
	MyModifica=True
	MyCancella=True
Case "7"
	MyAggiungi=True
	MyModifica=True
	MyCancella=True
Case Else

End SELECT

if (MyAggiungi OR MyModifica OR MyCancella) then
	MyAbsPath = Request.ServerVariables("PATH_INFO")
	if MyIDName="" then
		annulla=True
		MyErr = "Error: missing Primary Key Field"
		MyAggiungi = False
		MyModifica = False
		MyCancella = False
	end if
else
	MyIDName=""
end if
%>
<html>

<head>
<title><%Response.write(MyTitolo)%></title>

<meta name="Microsoft Border" content="none"></head>

<body style="font-family: Arial">

<p align="center"><img src="logo.gif" WIDTH="81" HEIGHT="78"><font face="Courier" color="#000080"><big><big><strong>
db Administration - <%Response.write(MyTitolo)%></strong></big></big></font><br>
</p>
<div align="center"><div align="center"><center>

<table border="0">
  <tr>
    <td><form method="POST" action="sql_aggiungi.asp">
      <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="RetSQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyAbsPath)%>"><input type="hidden" name="RetTitolo" value="<%Response.write(MyTitolo)%>"><input type="hidden" name="RetRitorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="RetComandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><p><%if MyAggiungi then%><input type="submit" value="Add" name="B1" style="font-size: 8pt; margin: 0px; padding: 0px"><%end if%></p>
    </form>
    </td>
    <td><form method="POST" action="<%Response.write(MyRitorno)%>">
      <input type="hidden" name="DSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="SubCommand" value="<%Response.write(MyMainSubCommand)%>"><p><%if MyRitorno<>"" then%><input type="submit" value="Back" name="B1" style="font-size: 8pt; margin: 0px; padding: 0px"><%end if%></p>
    </form>
    </td>
  </tr>
</table>
</center></div><%
' Interruzzione per input dati errato
if annulla then
	Response.Write("<CENTER><H4>ATTENTION</H4></CENTER>")
	Response.Write("<CENTER>" & MyErr & "</CENTER><br>")
else
	' apertura connessione
	objConn.open MyDSN,"",""

	' apertura dati db
	objRS.Open MySQL, objConn, 2, 1

	MyStart = 0
	While Not objRS.EOF
		MyStart = MyStart + 1
		objRS.MoveNext
	Wend

	objRS.MoveFirst

	if MyStart > 25 then
		MyPages = (Int(MyStart / 25)) * 25
		if MyStart > MyPages then
			MyPages = Int(MyStart / 25) + 1
		else
			MyPages = Int(MyStart / 25)
		end if
		MyStart = 1
		if Not IsNull(request.form("Pag")) then
			MyPage = Int(request.form("Pag"))
			if MyPage > MyPages then
				MyPage = MyPages
			elseif MyPage < 1 then
				MyPage = 1
			end if
		end if
	else
		MyStart = 1
		MyPage = 1
		MyPages = 1
	end if

	MyStart = ((MyPage - 1) * 25)
%>
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="right"></td>
    <td align="right"></td>
    <td align="right"><form method="POST" action="sql_table.asp">
      <input type="hidden" name="Comandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><input type="hidden" name="Pag" value="<%Response.write(MyPage-1)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="SQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><p><input type="submit" value="&lt;&lt;" name="Indietro" style="font-size: 10pt; background-color: rgb(255,255,255); text-decoration: underline; border: medium none rgb(255,255,255)">&nbsp;
      </p>
    </form>
    </td>
    <td style="font-size: 10pt; color: rgb(0,0,128); padding-top: 1px" valign="top" align="center">[page <%=MyPage%>/<%=MyPages%>]&nbsp;</td>
    <td><form method="POST" action="sql_table.asp">
      <input type="hidden" name="Comandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><input type="hidden" name="Pag" value="<%Response.write(MyPage+1)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="SQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><p><input type="submit" value="&gt;&gt;" name="Avanti" style="font-size: 10pt; background-color: rgb(255,255,255); text-decoration: underline; border: medium none rgb(255,255,255)">
      &nbsp;&nbsp; </p>
    </form>
    </td>
    <td align="center"><form method="POST" action="sql_table.asp">
      <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="SQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="Comandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><div align="right"><p><select name="Pag" size="1" style="font-size: 9pt; border: 0px none rgb(128,128,128); margin: 0px; padding-top: 3px">
<%For MyContinua = 1 to MyPages%>        <option value="<%Response.write(MyContinua)%>"><%Response.write(MyContinua)%></option>
<%Next%>      </select><input type="submit" value="Go" name="Vai" style="font-size: 8pt"></p>
      </div>
    </form>
    </td>
  </tr>
</table>
</center></div></div><div align="center"><%
	' controllo esistenza dati
	if Not objRS.EOF then
		objRS.MoveFirst
		' segue disegno tabella
		if (MyModifica OR MyCancella) then
%>
<div align="center"><center>

<table BORDER="1" COLS="<%=objRS.Fields.Count%>" height="53">
  <tr>
<%
			For each objField in objRS.Fields
				MyCmd = objField.Name
				if MyCmd=MyIDName then
					if (MyModifica AND MyCancella) then
%>
    <th colspan="2" height="16"><p align="center"><font face="Arial Narrow" size="2">Cmds</font></th>
<%					else%>
    <th height="16"><p align="center"><font face="Arial Narrow" size="2">Cmds</font></th>
<%					end if
				else%>
    <th height="16"><font face="Arial Narrow" size="2"><%=MyCmd%></font></th>
<%				End if%>
<%			Next%>
  </tr>
<%
			objRS.MoveFirst
			objRS.Move(MyStart)
			MyContinua = 1
			Do while (MyContinua < 26) AND (Not objRS.EOF) %>
  <tr>
<%
				For Each objField in objRS.Fields
					if IsNull(objField) then%>
    <td align="right" style="padding: 0px" valign="middle" height="25"><small><font face="Arial"><small><small><small><%Response.Write("&nbsp;")%></small></small></small></font></small></td>
<%					elseif objField.Name=MyIDName then
						if MyModifica then%>
    <td align="middle" height="25"><form method="POST" action="sql_modifica.asp" style="margin: 0px; padding: 0px">
      <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="RetSQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="IDValue" value="<%Response.write(objField.Value)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyAbsPath)%>"><input type="hidden" name="RetTitolo" value="<%Response.write(MyTitolo)%>"><input type="hidden" name="RetRitorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="RetComandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><p><font face="Arial" size="1"><input type="submit" value="Modify" name="Modifica" style="font-size: 8pt; margin: 0px; padding: 0px"> </font></p>
    </form>
    </td>
<%						end if%>
<%						if MyCancella then%>
    <td align="middle" height="25"><form method="POST" action="sql_elimina.asp" style="margin: 0px; padding: 0px">
      <input type="hidden" name="DSN" value="<%Response.write(MyDSN)%>"><input type="hidden" name="Table" value="<%Response.write(MyTable)%>"><input type="hidden" name="RetSQL" value="<%Response.write(MySQL)%>"><input type="hidden" name="IDName" value="<%Response.write(MyIDName)%>"><input type="hidden" name="IDValue" value="<%Response.write(objField.Value)%>"><input type="hidden" name="Titolo" value="<%Response.Write(MyTitolo)%>"><input type="hidden" name="Ritorno" value="<%Response.write(MyAbsPath)%>"><input type="hidden" name="RetTitolo" value="<%Response.write(MyTitolo)%>"><input type="hidden" name="RetRitorno" value="<%Response.write(MyRitorno)%>"><input type="hidden" name="RetComandi" value="<%Response.write(MyModCmd)%>"><input type="hidden" name="MainDSN" value="<%Response.write(MyMainDSN)%>"><input type="hidden" name="MainRitorno" value="<%Response.write(MyMainRitorno)%>"><input type="hidden" name="MainSubCommand" value="<%Response.write(MyMainSubCommand)%>"><p><font face="Arial" size="1"><input type="submit" value="Delete" name="Elimina" style="font-size: 8pt; margin: 0px; padding: 0px"> </font></p>
    </form>
    </td>
<%						end if
					else%>
    <td align="right" style="padding: 0px" valign="middle" height="25"><small><font face="Arial"><%	if IsNull(objField) then
							Response.Write("&nbsp;")
						else
							MyObjField.SetField = objField
							MyObjField.SetSmartRef=True
							Response.Write(MyObjField.NewParser)
						end if%></font></small></td>
<%					end if%>
<%				Next
				objRS.MoveNext%>
  </tr>
<%				MyContinua = MyContinua + 1
			Loop%>
</table>
</center></div></div><%
		else%>
<div align="center"><div align="center"><center>

<table BORDER="1" COLS="<%=objRS.Fields.Count%>">
  <tr>
<% 			For each objField in objRS.Fields %>
    <th><font face="Arial Narrow"><font size="2"><%=objField.Name%></font><small> </small></font></th>
<%			Next%>
  </tr>
<%
			objRS.MoveFirst
			objRS.Move(MyStart)
			MyContinua = 1
			Do while (MyContinua < 26) AND (Not objRS.EOF) %>
  <tr>
<% 				For Each objField in objRS.Fields %>
    <td align="right"><p align="left"><small><%
					if IsNull(objField) then
						Response.Write("&nbsp;")
					else
						Response.Write(objField.Value)
					end if%></small></td>
<%				Next
				objRS.MoveNext%>
  </tr>
<%				MyContinua = MyContinua + 1
			Loop%>
</table>
</center></div>

<p><%	end if
		' chiusura
		objRS.close
	end if
end if
%></p>

<p>&nbsp;<%
if MyIDName<>"" then
	Response.Write("<br>")
end if
%> </p>

<p align="center"><font color="#000080" face="Arial Narrow"><strong><%
if MyRitorno="" then
	Response.write("<br>Close the window")
end if%></strong></font></p>
</div>

<p align="center"><%
Set annulla=Nothing
Set MyErr=Nothing

Set MyDSN=Nothing
Set MyTable=Nothing
Set MySQL=Nothing
Set MyIDName=Nothing
Set MyTitolo=Nothing
Set MyRitorno=Nothing
Set MyModCmd=Nothing
Set MyAggiungi=Nothing
Set MyModifica=Nothing
Set MyCancella=Nothing
Set MyAbsPath=Nothing
Set MyCmd=Nothing
Set MyMainDSN=Nothing
Set MyMainSubCommand=Nothing
Set MyMainRitorno=Nothing
Set MyObjField=Nothing
Set MyStart = Nothing
Set MyPage = Nothing
Set MyPages = Nothing
Set MyContinua = Nothing

Set objRS=Nothing
Set objConn=Nothing

Session.Abandon
%> </p>

<p>&nbsp; </p>
</body>
</html>
