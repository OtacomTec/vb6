<%@ Language=VBScript %>
<%
'--------------------------------------------------------------------------------------------
' Modulo Inserimento Dati
' G.B. Tassara - 06/2002
'--------------------------------------------------------------------------------------------
' Parametri in ingresso
'
' DSN			testo - nome connessione DSN
' Table		testo - nome tabella - il primo campo deve essere l'identificativo del record
' SQL			testo - espressione SQL - se nulla viene eseguita una query di accesso
'						 completo alla tabella MyTable
' Titolo		testo - titolo pagina - se omesso verra' utilizzata la dicitura Dati db
'
'--------------------------------------------------------------------------------------------

Response.Buffer=True

' classe esecuzione SQL
Dim MyRunSQL
Set MyRunSQL = New clsRunSQL

' classe caratteristiche Field
Dim MyObjField
Set MyObjField = New clsObjField

' connessioni
Dim objConn, objRS

' flag valore nullo
Dim MyNull

' flag annullamento operazione
Dim annulla
annulla=False

' codice errore
Dim MyErr

' variabili di ingresso
Dim MyDSN, MySQL, MySQLCmd, MySQLValue, MyTable, MyIDName, MyIDValue, MyTitolo

' lettura valori DSN e tabella e controllo degli errori
' 1 - nome DSN
MyDSN = cstr(request.form("DSN"))

' 2 - nome Tabella
MyTable = cstr(request.form("Table"))

' 3 - titolo
MyTitolo = cstr(request.form("Titolo"))

' 4 - ID Name
MyIDName = cstr(request.form("IDName"))

' 5 - ID Value
MyIDValue = cstr(request.form("IDValue"))

' 6 - nome SQL
MySQL = cstr(request.form("SQL"))

' controlli sui campi di ingresso obbligatori
if MyDSN="" then
	annulla=True
	MyErr = "Error: missing DSN"
end if

if MyTable="" then
	annulla=True
	MyErr = "Error: missing Table Name"
end if

if MySQL="" then
	annulla=True
	MyErr = "Error: missing SQL query"
elseif MyTable="" or MyIDValue= "" or MyIDName="" then
	annulla=True
	MyErr = "Error: missing Table Name"
end if

' controllo sul titolo
if MyTitolo="" then
	MyTitolo = "Confirm dB Data Operation"
else
	MyTitolo = "Confirm " & MyTitolo
end if
%>
<!-- #include file="sql_cmd.cls" -->
<!-- #include file="sql_field.cls" -->
<%
' costruzione istruzione SQL
' fase 1 - apertura tabella per prelevamento nome campi e caratteristiche
Set objRs=Server.CreateObject("ADODB.Recordset")
Set objConn=Server.CreateObject("ADODB.Connection")

Response.write(MyDSN)

objConn.open MyDSN,"",""
objRS.Open MySQL, objConn, 2, 1

' preparazione istruzione SQL
Divisore = ""
MySQLCmd = "UPDATE " & MyTable & " SET "
MySQLValue = ""

' loop su tutti i campi della tabella
For Each objField In objRS.Fields

	Nome = objField.Name
	Valore = request.form(Nome)

	' trattamento valore nullo
	MyNull = FALSE
	If IsNull(Valore) then
		MyNull = TRUE
		Valore = "Null"
	elseif IsEmpty(Valore) then
		MyNull = TRUE
		Valore = "Null"
	elseif Len(Valore)<1 then
		MyNull = TRUE
		Valore = "Null"
	end if

	if Not IsNull(Valore) then

		' controllo equivalente di stringa
		MyStringa = "'" & Valore & "'"
		MyLen = Len(MyStringa)

		If MyLen > 2 then

			' imposto i valori per la classe
			MyObjField.SetField = objField
			' imposto il valore del campo
			MyObjField.SetValue = Valore

			Verifica = MyObjField.Updatable

			' calcolo il valore di ritorno con i separatori inclusi		
			if MyObjField.Updatable then
					'estraggo il valore
					if MyNull then
						Valore = "Null"
					else
						Valore = MyObjField.NewValue
					end if

					'update elenco valori SQL
					MySQLValue = MySQLValue & Divisore & " " & Nome & " = " & Valore

					if Divisore <> "," then
						Divisore = ","
					end if
			end if
		end if
	end if
Next

MySQLCmd = MySQLCmd & MySQLValue & " WHERE (((" & MyTable & "." & MyIDName & ")=" & MyIDValue & "));"
objRS.close

Set objRS=Nothing
Set objConn=Nothing
%>
<html>

<head>
<title><%Response.write(MyTitolo)%></title>
<meta name="Microsoft Border" content="none">
</head>

<body style="font-family: Arial">

<p align="center"><img src="logo.gif" width="81" height="78"><font face="Courier"
color="#000080"><big><big><strong>&nbsp; <%Response.write(MyTitolo)%></strong></big></big></font><br>
</p>

<p align="center"><%
' Interruzzione per input dati errato
if annulla then
	Response.Write("<CENTER><H4>ATTENTION</H4></CENTER>")
	Response.Write("<CENTER>" & MyErr & "</CENTER><br>")
else
	MyRunSQL.DSN = MyDSN
	MyRunSQL.SQL = MySQLCmd
	MyErr = MyRunSQL.Result
	Response.Write(MyErr)
end if

Set annulla=Nothing
Set MyErr=Nothing

Set MyDSN=Nothing
Set MySQL=Nothing
Set MySQLCmd=Nothing
Set MySQLValue=Nothing
Set MyTable=Nothing
Set MyIDName=Nothing
Set MyIDValue=Nothing
Set MyTitolo=Nothing
Set MyNull = Nothing

Set MyRunSQL=Nothing
Set MyObjField = Nothing

Session.Abandon
%></p>

<p align="center"><font color="#000080" face="Arial Narrow"><big>Close the window</big> </font></p>
</body>
</html>
