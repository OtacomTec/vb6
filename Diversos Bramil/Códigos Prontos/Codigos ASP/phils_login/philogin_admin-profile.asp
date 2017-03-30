<%
Option Explicit
Dim sql, rsProfile, username, malechecked, femalechecked, sendemailchecked

username = Request.QueryString("username")
%>

<html>
<head>
<meta http-equiv="Expires" content="Mon, 06 Jan 1990 00:00:01 GMT">
<title><%=username%>'s Profile</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size="1">
<%if Request.Cookies("philoginadmin") = "True" then%>

<!--#include file="conn.asp"-->
<%
sql = "SELECT * FROM Users WHERE username = '" & username & "'"
Set rsProfile = Server.CreateObject("ADODB.Recordset")
rsProfile.Open sql, conn, 3, 3
%>

<img src="icons/<%=rsProfile("icon")%>.gif" align="right">

<h3><%=username%>'s perfil</h3>

<%if Request.QueryString("updated") = "true" then%>
<p><font color="#DD0000">O perfil</font><font color="#DD0000"> foi alterado com 
  sucesso <%=Time()%></font></p>
<%else%>
<br clear="all">
<%end if%>

<form action="philogin_admin-profile_update.asp" method="post" name="profileform">
<input type="hidden" name="username" value="<%=username%>">
<table width="100%" cellpadding="2" cellspacing="0" border="0">
<tr>
      <td><font face="arial,helvetica" size="1">Senha :</font></td>
  <td><font face="arial,helvetica" size="1"><input type="password" name="password" size="20" value="<%=rsProfile("password")%>"></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Nome :</font></td>
  <td><font face="arial,helvetica" size="1"><input type="text" name="firstname" size="20" value="<%=rsProfile("firstname")%>"></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Sobrenome :</font></td>
  <td><font face="arial,helvetica" size="1"><input type="text" name="surname" size="20" value="<%=rsProfile("surname")%>"></font></td>
</tr>
<tr>
  <td><font face="arial,helvetica" size="1">Email :</font></td>
  <td><font face="arial,helvetica" size="1"><input type="text" name="email" size="20" value="<%=rsProfile("email")%>"></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Nasc. (dd/mm/yyyy) :</font></td>
  <td><font face="arial,helvetica" size="1"><input type="text" name="birth_day" value="<%=day(rsProfile("dob"))%>" size="2" maxlength="2"> <input type="text" name="birth_month" value="<%=month(rsProfile("dob"))%>" size="2" maxlength="2"> <input type="text" name="birth_year" value="<%=year(rsProfile("dob"))%>" size="4" maxlength="4"></font></td>
</tr>
<tr>
  <td><font face="arial,helvetica" size="1">Starsign :</font></td>
  <td><font face="arial,helvetica" size="1"><%=rsProfile("starsign")%></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Sexo :</font></td>
  <td><font face="arial,helvetica" size="1"><%
  if rsProfile("sex") = "male" then
    malechecked = " checked"
  else
    femalechecked = " checked"
  end if
  %><input type="radio" name="sex" value="male"<%=malechecked%>>male
  <input type="radio" name="sex" value="female"<%=femalechecked%>>female</font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size=1>Receber altera&ccedil;&otilde;es 
        por e-mail:</font></td>
  <td><font face="arial,helvetica" size=1><%
  if rsProfile("sendemail") = True then
    sendemailchecked = " checked"
  else
    sendemailchecked = ""
  end if
  %><input type="checkbox" name="sendemail" value="True"<%=sendemailchecked%>></font></td>
</tr>
<tr>
  <td colspan=2 align="center"><font face="arial,helvetica" size="1"><input type="submit" name="submitbutton" value="Mudar detalhes" onClick="profileform.submitbutton.value='Please wait...'"></font></td>
</tr>
</table>

<center>
    <b><a href="javascript:window.close()">Fechar esta janela</a></b> 
  </center>

</form>

<%
rsProfile.close
set rsProfile = nothing
conn.close
set conn = nothing
%>

<%else%>
<p>V&aacute; embora, vc n&atilde;o est&aacute; logado como administrador. </p>
<%end if%>
</font>
</body>
</html>
