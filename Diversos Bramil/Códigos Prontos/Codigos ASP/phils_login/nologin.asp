<html>
<head>
<title>Must be logged in to view this page</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2 color="#FFFFFF">

<h3 align="center">The page you were trying to view is only available to registered users, please log in or sign up!</h3>

<form name="signin" action="signin.asp" method="post">
<input type="hidden" name="page" value="<%=Request.QueryString("page")%>">
<table width="420" cellpadding=2 cellspacing=0 align="center" border=1 bordercolor="#000000">
<tr>
      <th bgcolor="#000000"><font face="arial,helvetica" size=2 color="#FFFFFF">Login</font></th>
      <th bgcolor="#000000"><font face="arial,helvetica" size="2" color="#FFFFFF">Cadastro</font></th>
</tr>
<tr>
      <td align="center"> <font face="arial,helvetica" size=2><b> Usu&aacute;rio 
        : 
        <input type="text" name="username" size="12" style="font-family: tahoma,verdana, Arial, Helvetica, sans-serif;"><br>
        Senha : 
        <input type="password" name="password" size="12" style="font-family: tahoma,verdana, Arial, Helvetica, sans-serif;"><br>
        Lembrar senha: 
        <input type="checkbox" name="stayloggedin" value="yes"><br>
  <input type="submit" value="Login" style="font-family: tahoma,verdana, Arial, Helvetica, sans-serif;">
  </b></font>
  </td>
      <td align="center"> <font face="arial,helvetica" size=2> </font><font face="arial,helvetica" size=1><b><a href="signupform.asp"><font size="2">Ainda 
        n&atilde;o sou membro<br>
        inscreva-me agora!</font></a></b></font><font face="arial,helvetica" size=2><a href="signupform.asp"></a> 
        </font> </td>
</tr>
</table>
</form>

  </font>
  </td>
</tr>
</table>

</body>
</html>
