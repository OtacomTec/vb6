<%
Option Explicit
Dim sql, rsProfile, username, malechecked, femalechecked, sendemailchecked

username = Request.Cookies("username")
%>

<html>
<head>
<meta http-equiv="Expires" content="Mon, 06 Jan 1990 00:00:01 GMT">
<title>Your Profile</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size="1">
<%
if username = "" then
  nologin()
end if
%>
<!--#include file="conn.asp"-->
<%
sql = "SELECT * FROM Users WHERE username = '" & username & "'"
Set rsProfile = Server.CreateObject("ADODB.Recordset")
rsProfile.Open sql, conn, 3, 3
%>

<img src="icons/<%=rsProfile("icon")%>.gif" align="right">

<h3><%=username%>'s perfil</h3>

<%if Request.QueryString("updated") = "true" then%>
<p><font color="#DD0000">Seu perfil foi alterado com sucesso </font><font color="#DD0000"><%=Time()%></font></p>
<%else%>
<br clear="all">
<%end if%>

<form action="profile_update.asp" method="post" name="profileform">
<table width="100%" cellpadding="2" cellspacing="0" border="0">
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
      <td><font face="arial,helvetica" size="1">Receber novidades por e-mail:</font></td>
  <td><font face="arial,helvetica" size=1><%
  if rsProfile("sendemail") = True then
    sendemailchecked = " checked"
  else
    sendemailchecked = ""
  end if
  %><input type="checkbox" name="sendemail" value="True"<%=sendemailchecked%>></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Entrar senha:</font></td>
  <td><font face="arial,helvetica" size="1"><input type="password" name="password" size="20"></font></td>
</tr>
<tr>
  <td colspan=2 align="center"><font face="arial,helvetica" size="1"><input type="submit" name="submitbutton" value="Mudar detalhes" onClick="profileform.submitbutton.value='Please wait...'"></font></td>
</tr>
</table>
<center>
    <b><a href="changepwd.asp?icon=<%=rsProfile("icon")%>">Mudar senha</a><br>
    <a href="javascript:window.close()">Fechar esta janela</a></b> 
  </center>

</form>

</font>
</body>
</html>

<%
rsProfile.close
set rsProfile = nothing
conn.close
set conn = nothing
%>

<%Function nologin()%>

<p align="center"><b>Voc&ecirc; precisa estar logado para mudar seu perfil.</b></p>

<p align="center"><b><a href="javascript:window.close()">Fechar esta janela</a></b></p>

</font>
</body>
</html>
<%Response.end
End Function%>