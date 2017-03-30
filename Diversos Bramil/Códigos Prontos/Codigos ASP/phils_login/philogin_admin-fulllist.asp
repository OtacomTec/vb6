<%
Option Explicit
Dim sql,rsUsers

if Request.Cookies("philoginadmin") = "True" then
%>
<!--#include file="conn.asp"-->
<%
sql = "SELECT * FROM users ORDER BY username"
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

<p><a href="philogin_admin-signout.asp">Sair do sistema</a></p>
  
<%if not rsUsers.EOF then%>
  
<table cellpadding="2" cellspacing="0" border="1" bordercolor="#B4B4B4" width="100%">
<tr bgcolor="#B4B4B4">
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Membro</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Senha</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Nome</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Sobrenome</font></th>
  <th><font face="arial,helvetica" size="2" color="#FFFFFF">Email</font></th>
  <th><font face="arial,helvetica" size="2" color="#FFFFFF">DOB</font></th>
  <th><font face="arial,helvetica" size="2" color="#FFFFFF">Starsign</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Sexo</font></th>
    <th><font face="arial,helvetica" size="2" color="#FFFFFF">Icone</font></th>
</tr>
<%rsUsers.Movefirst
do until rsUsers.EOF%>
<tr>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("username")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("password")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("firstname")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("surname")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("email")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("dob")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("starsign")%></font></td>
  <td><font face="arial,helvetica" size="2"><%=rsUsers("sex")%></font></td>
  <td align="center"><img src="icons/<%=rsUsers("icon")%>_small.gif" width="20" height="20"></td>
</tr>
<%rsUsers.Movenext
loop%>
</table>

<%else%>

<p><b>Desculpe, sem membros cadastrados.</b></p>

<%end if%>

<%else
Response.Redirect("philogin_admin.asp")
end if%>

</font>
</body>
</html>

<%
rsUsers.close
set rsUsers = nothing
conn.close
set conn = nothing
%>
