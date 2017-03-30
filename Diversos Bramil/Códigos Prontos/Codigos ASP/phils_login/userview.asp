<%
Option Explicit
Dim sql,rsUsers,rsUser,username,startletter,alphabet

username = Request.Cookies("username")
startletter = Request.QueryString("startletter")

'Example of protecting a page for members only
if username = "" then
  Response.Redirect("nologin.asp?page=userview.asp")
end if

'If no letter has been clicked (this is 1st visit), start displaying at letter A
if startletter = "" then
  startletter = "A"
end if

%>
<!--#include file="conn.asp"-->
<%

'If 123 option is chosen, pick all records that don't start with a letter (this makes a BIG sql string)
if startletter = "nonalphabet" then
  sql = "SELECT username, icon FROM users WHERE username Not Like 'a%'"
  for alphabet = 98 to 122
	sql = sql & " AND username Not Like '" & chr(alphabet) & "%'"
  next
  sql = sql & " ORDER BY username"
'Otherwise just get the users that start with the chosen letter
else
  sql = "SELECT username, icon FROM users WHERE username Like '" & startletter & "%' ORDER BY username"
end if
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open sql, conn, 3, 3

sql = "SELECT icon FROM Users WHERE username = '" & username & "'"
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open sql, conn, 3, 3
%>

<html>
<head>
<title>User List</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2>
    
<h3>Lista de membros</h3>
  
<p>aqui est&aacute; a lista com todos os membros cadastrados. Clique nas letras 
  correspondentes aos nomes dos membros.Voc&ecirc; poder&aacute; mandar uma mensagem 
  para qualquer um deles clicando no envelope.</p>

<!-- Displays A-Z -->
<table cellpadding="2" cellspacing="0" border="1" bordercolor="#B4B4B4" width="468">
  <tr>
	<td align="center" bgcolor="#FFFFFF"><font face="arial,helvetica" size="2"><b><a href="userview.asp?startletter=nonalphabet">123</a></b></font>
<%'Loop through alphabet (chars 65 to 90 are A-Z)
for alphabet = 65 to 90%>
    <td align="center" bgcolor="#FFFFFF"><font face="arial,helvetica" size="2"><b><a href="userview.asp?startletter=<%=chr(alphabet)%>"><%=chr(alphabet)%></a></b></font></td>
<%next%>
  </tr>
  </table>
  
<%if not rsUsers.EOF then%>

<!-- Displays users -->
  <table cellpadding="2" cellspacing="0" border="1" bordercolor="#B4B4B4" width="468">
  <tr bgcolor="#B4B4B4">
	<th><font face="arial,helvetica" size="2" color="#FFFFFF">membro</font></th>
	<th><font face="arial,helvetica" size="2" color="#FFFFFF">Icone</font></th>
	<th><font face="arial,helvetica" size="2" color="#FFFFFF">Contato</font></th>
  </tr>
  <%rsUsers.Movefirst
  do until rsUsers.EOF%>
  <tr>
    <td><font face="arial,helvetica" size="2"><%=rsUsers("username")%></font></td>
	<td align="center"><img src="icons/<%=rsUsers("icon")%>_small.gif" width="20" height="20"></td>
	<td align="center"><font face="arial,helvetica" size="1"><a href="messagecompose.asp?senduser=<%=rsUsers("username")%>"><img src="icons/envelope.gif" alt="send a message to <%=rsUsers("username")%>" border=0 hspace="10">enviar 
      mensagem </a></font></td>
  </tr>
  <%rsUsers.Movenext
  loop%>
  </table>

<%'If no users, give a message
else%>

<p><b>desculpe, n&atilde;o existem membros com esta letra <%=startletter%>.</b></p>

<%end if%>

<br>
<p><a href="index.asp">Voltar &aacute; p&aacute;gina inicial</a></p>

</font>
</body>
</html>

<%
rsUsers.close
set rsUsers = nothing
rsUser.close
set rsUser = nothing
conn.close
set conn = nothing
%>
