<%
Option Explicit
Dim username, icon

username = Request.Cookies("username")
icon = Request.QueryString("icon")
%>

<html>
<head>
<title>troca de senha</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size="1">
<%
if username = "" then
  nologin()
end if
%>

<img src="icons/<%=icon%>.gif" align="right">

<h3>Troca de senha</h3>

<%if Request.QueryString("updated") = "true" then%>
<p><font color="#DD0000">Sua senha foi trocada com sucesso </font><font color="#DD0000"><%=Time()%></font></p>
<%else%>
<br clear="all">
<%end if%>

<form action="pwd_update.asp" method="post" name="passwordform">
<input type="hidden" name="icon" value="<%=icon%>">
<table width="100%" cellpadding="2" cellspacing="0" border="0">
<tr>
      <td><font face="arial,helvetica" size="1">Senha antiga:</font></td>
  <td><font face="arial,helvetica" size="1"><input type="password" name="oldpassword" size="20"></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Nova senha:</font></td>
  <td><font face="arial,helvetica" size="1"><input type="password" name="newpassword" size="20"></font></td>
</tr>
<tr>
      <td><font face="arial,helvetica" size="1">Confirme a nova senha:</font></td>
  <td><font face="arial,helvetica" size="1"><input type="password" name="newpasswordconfirm" size="20"></font></td>
</tr>
<tr>
  <td colspan=2 align="center"><font face="arial,helvetica" size="1"><input type="submit" name="submitbutton" value="Mude a senha" onClick="passwordform.submitbutton.value='Please wait...'"></font></td>
</tr>
</table>
<center>
    <p><b><a href="profile.asp">voltar ao perfil</a></b></p>
    <p><b>
      <a href="javascript:window.close()">fechar esta janela</a></b> </p>
  </center>

</form>

</font>
</body>
</html>

<%Function nologin()%>

<p align="center"><b>Voc&ecirc; precisa estar logado para ver o seu perfil.</b></p>

<p align="center"><b><a href="javascript:window.close()">fechar esta janela</a></b></p>

</font>
</body>
</html>
<%Response.end
End Function%>