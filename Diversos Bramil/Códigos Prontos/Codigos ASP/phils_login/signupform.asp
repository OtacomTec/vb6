<%
Option Explicit
Dim username

username = Request.Cookies("username")
%>

<html>
<head>
<title>Signup Form</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2>

<%'See if they're actually already logged in
if username <> "" then%>

<p><b>Voc&ecirc; est&aacute; logado.</b></p>

<p>Se voc&ecirc; quer preencher um novo cadastro, precisa <a href="signout.asp"><b>sair 
  do sistema </b></a>primeiro.</p>

<%'Otherwise display the sign up form
else%>

<form name="signup" action="signupprocess.asp" method="post">
<table width="420" cellpadding=2 cellspacing=0 align="center">
<tr>
      <th colspan=2 bgcolor="#000000"><font face="arial,helvetica" size="2" color="#FFFFFF">Cadastro 
        de membros</font></td> 
    </tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Nome</b></font></td>
  <td align="center"><input type="text" name="firstname" size=13></td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Sobrenome</b></font></td>
  <td align="center"><input type="text" name="surname" size=13></td>
</tr>
<tr>
  <td align="center"><font face="arial,helvetica" size=2><b>Email</b></font></td>
  <td align="center"><input type="text" name="email" size=13></td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Data de nascimento</b></font></td>
  <td align="center">
  <input type="text" name="birth_day" value="dd" size="2" maxlength="2"> <input type="text" name="birth_month" value="mm" size="2" maxlength="2"> <input type="text" name="birth_year" value="yyyy" size="4" maxlength="4">
  </td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Sexo</b></font></td>
  <td align="center"><font face="arial,helvetica" size=2><b><input type="radio" name="sex" value="male">
        Homem 
        <input type="radio" name="sex" value="female">
        Mulher</b></font></td>
</tr>
<tr>
  <td align="center" colspan=2><hr color="#000000"></td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Usu&aacute;rio</b></font></td>
  <td align="center"><input type="text" name="username" size=13></td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Senha</b></font></td>
  <td align="center"><input type="password" name="password" size=13></td>
</tr>
<tr>
      <td align="center"><font face="arial,helvetica" size=2><b>Confirme sua senha</b></font></td>
  <td align="center"><input type="password" name="passwordconfirm" size=13></td>
</tr>
<tr>
      <td align="center" colspan=2><font face="arial,helvetica" size=1><b>(note: 
        sua senha deve ter entre 5 e 15 caracteres)</b></font></td>
</tr>
<tr>
  <td align="center" colspan=2><hr color="#000000"></td>
</tr>
<tr>
      <td align="center" colspan=2> <font face="arial,helvetica" size=2><b>Escolha 
        uma imagem para voc&ecirc;:</b></font><br>
        <br>
  <table width="100%" cellpadding=2 cellspacing=0>
  <tr>
	<td align="center"><img src="icons/1.gif" onClick="document.signup.icon[0].checked=true"><br><input type="radio" name="icon" value="1"></td>
	<td align="center"><img src="icons/2.gif" onClick="document.signup.icon[1].checked=true"><br><input type="radio" name="icon" value="2"></td>
	<td align="center"><img src="icons/3.gif" onClick="document.signup.icon[2].checked=true"><br><input type="radio" name="icon" value="3"></td>
	<td align="center"><img src="icons/4.gif" onClick="document.signup.icon[3].checked=true"><br><input type="radio" name="icon" value="4"></td>
	<td align="center"><img src="icons/5.gif" onClick="document.signup.icon[4].checked=true"><br><input type="radio" name="icon" value="5"></td>
  </tr>
  <tr>
	<td align="center"><img src="icons/6.gif" onClick="document.signup.icon[5].checked=true"><br><input type="radio" name="icon" value="6"></td>
	<td align="center"><img src="icons/7.gif" onClick="document.signup.icon[6].checked=true"><br><input type="radio" name="icon" value="7"></td>
	<td align="center"><img src="icons/8.gif" onClick="document.signup.icon[7].checked=true"><br><input type="radio" name="icon" value="8"></td>
	<td align="center"><img src="icons/9.gif" onClick="document.signup.icon[8].checked=true"><br><input type="radio" name="icon" value="9"></td>
	<td align="center"><img src="icons/10.gif" onClick="document.signup.icon[9].checked=true"><br><input type="radio" name="icon" value="10"></td>
  </tr>
  <tr>
	<td align="center"><img src="icons/11.gif" onClick="document.signup.icon[10].checked=true"><br><input type="radio" name="icon" value="11"></td>
	<td align="center"><img src="icons/12.gif" onClick="document.signup.icon[11].checked=true"><br><input type="radio" name="icon" value="12"></td>
	<td align="center"><img src="icons/13.gif" onClick="document.signup.icon[12].checked=true"><br><input type="radio" name="icon" value="13"></td>
	<td align="center"><img src="icons/14.gif" onClick="document.signup.icon[13].checked=true"><br><input type="radio" name="icon" value="14"></td>
	<td align="center"><img src="icons/15.gif" onClick="document.signup.icon[14].checked=true"><br><input type="radio" name="icon" value="15"></td>
  </tr>
  <tr>
	<td align="center"><img src="icons/16.gif" onClick="document.signup.icon[15].checked=true"><br><input type="radio" name="icon" value="16"></td>
	<td align="center"><img src="icons/17.gif" onClick="document.signup.icon[16].checked=true"><br><input type="radio" name="icon" value="17"></td>
	<td align="center"><img src="icons/18.gif" onClick="document.signup.icon[17].checked=true"><br><input type="radio" name="icon" value="18"></td>
	<td align="center"><img src="icons/19.gif" onClick="document.signup.icon[18].checked=true"><br><input type="radio" name="icon" value="19"></td>
	<td align="center"><img src="icons/20.gif" onClick="document.signup.icon[19].checked=true"><br><input type="radio" name="icon" value="20"></td>
  </tr>
  <tr>
	<td align="center"><img src="icons/21.gif" onClick="document.signup.icon[20].checked=true"><br><input type="radio" name="icon" value="21"></td>
	<td align="center"><img src="icons/22.gif" onClick="document.signup.icon[21].checked=true"><br><input type="radio" name="icon" value="22"></td>
	<td align="center"><img src="icons/23.gif" onClick="document.signup.icon[22].checked=true"><br><input type="radio" name="icon" value="23"></td>
	<td align="center"><img src="icons/24.gif" onClick="document.signup.icon[23].checked=true"><br><input type="radio" name="icon" value="24"></td>
	<td align="center"><img src="icons/25.gif" onClick="document.signup.icon[24].checked=true"><br><input type="radio" name="icon" value="25"></td>
  </tr>
  </table>
  
  </td>
</tr>
<tr>
	  <td align="center" colspan=2><font face="arial,helvetica" size="2">Envie-me 
        novidade por e-mail&nbsp; 
        <input type="checkbox" name="sendemail" checked></font></td>
</tr>
<tr>
  <td align="center" colspan=2><input type="submit" value="Cadastre-se"></td>
</tr>
</table>

</form>

<%end if%>

</font>
</body>
</html>
