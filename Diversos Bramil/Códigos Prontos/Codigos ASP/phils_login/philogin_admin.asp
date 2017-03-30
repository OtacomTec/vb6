<%
Option Explicit
Dim sql,rsUsers,username,startletter,alphabet,newcount

startletter = Request.QueryString("startletter")

if startletter = "" then
  startletter = "A"
end if
%>
<!--#include file="conn.asp"-->
<%
if startletter = "nonalphabet" then
  sql = "SELECT username, icon FROM users WHERE username Not Like 'a%'"
  for alphabet = 98 to 122
	sql = sql & " AND username Not Like '" & chr(alphabet) & "%'"
  next
  sql = sql & " ORDER BY username"
else
  sql = "SELECT username, icon FROM users WHERE username Like '" & startletter & "%' ORDER BY username"
end if
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open sql, conn, 3, 3
%>

<html>
<head>

<script language="JavaScript">
<!-- hide on

function popup(popupfile,winheight,winwidth,scrolls)
{
open(popupfile,"PopupWindow","resizable=no,height=" + winheight + ",width=" + winwidth + ",scrollbars=no" + scrolls);
}

// hide off -->
</script>

<title>Philogin User List</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size="2">

<h3> Administra&ccedil;&atilde;o</h3>

<%if Request.Cookies("philoginadmin") = "True" then%>

<p><a href="philogin_admin-signout.asp">Sair do sistema</a></p>

<table cellpadding="2" cellspacing="0" border="1" bordercolor="#B4B4B4" width="500">
<tr>
  <td align="center" bgcolor="#FFFFFF" id="cell123"><font face="arial,helvetica" size="2"><b><a href="philogin_admin.asp?startletter=nonalphabet" onMouseOver="cell123.bgColor='#B4B4B4'" onMouseOut="cell123.bgColor='#FFFFFF'">123</a></b></font>
<%'Loop through alphabet (chars 65 to 90 are A-Z)
for alphabet = 65 to 90%>
  <td align="center" bgcolor="#FFFFFF" id="cell<%=chr(alphabet)%>"><font face="arial,helvetica" size="2"><b><a href="philogin_admin.asp?startletter=<%=chr(alphabet)%>" onMouseOver="cell<%=chr(alphabet)%>.bgColor='#B4B4B4'" onMouseOut="cell<%=chr(alphabet)%>.bgColor='#FFFFFF'"><%=chr(alphabet)%></a></b></font></td>
<%next%>
</tr>
</table>
  
<%if not rsUsers.EOF then%>
  
<table cellpadding="2" cellspacing="0" border="1" bordercolor="#B4B4B4" width="500">
<tr bgcolor="#B4B4B4">
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Membro</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Icone</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Perfil</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Apagar membro</font></th>
</tr>
<%rsUsers.Movefirst
do until rsUsers.EOF%>
<tr>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("username")%></font></td>
  <td align="center"><img src="icons/<%=rsUsers("icon")%>_small.gif" width="20" height="20"></td>
    <td align="center"><font face="arial,helvetica" size="1"><a href="javascript:popup('philogin_admin-profile.asp?username=<%=rsUsers("username")%>',350,275,'no')">Ver/editar 
      perfil </a></font></td>
    <td align="center"><font face="arial,helvetica" size="1"><a href="philogin_admin-delete.asp?username=<%=rsUsers("username")%>&startletter=<%=startletter%>">Apagar 
      usu&aacute;rio</a></font></td>
</tr>
<%rsUsers.Movenext
loop%>
</table>

<%else%>

<p><b>Desculpe, n&atilde;o existem membros com esta letra<%=startletter%>.</b></p>

<%end if%>

<%else%>

<p><b>Voc&ecirc; n&atilde;o est&aacute; logado como administrador, por favor logue-se 
  abaixo:</b></p>

<form action="philogin_admin-login.asp" method="post">
<table cellpadding=2 cellspacing=0 border=0>
<tr>
      <td><font face="arial,helvetica" size="2"><b>Usu&aacute;rio</b></font></td>
  <td><input type="text" name="username" size="20"></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="2"><b>Senha</b></font></td>
  <td><input type="password" name="password" size="20"></td>
</tr>
<tr>
  <td colspan="2"><input type="submit" value="Log In"></td>
</tr>
</table>
</form>

<%end if%>

</font>
</body>
</html>

<%
rsUsers.close
set rsUsers = nothing
conn.close
set conn = nothing
%>
