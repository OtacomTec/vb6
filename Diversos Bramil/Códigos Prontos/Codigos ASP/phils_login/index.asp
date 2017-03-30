<%
Option Explicit
Dim sql,username,rsUser,rsMessages,newcount

username = Request.Cookies("username")

'If the username cookie is set, they must have logged in, so get their details from the database
if username <> "" then
%>
<!--#include file="conn.asp"-->
<%
  sql = "SELECT icon FROM Users WHERE username = '" & username & "'"
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open sql, conn, 3, 3
  
  sql = "SELECT messageread FROM messages WHERE sendto = '" & username & "'"
  Set rsMessages = Server.CreateObject("ADODB.Recordset")
  rsMessages.Open sql, conn, 3, 3
  
  newcount = 0
  if not rsMessages.EOF then
	rsMessages.Movefirst
	do until rsMessages.EOF
		if rsMessages("messageread") = False then
			newcount = newcount + 1
		end if
		rsMessages.Movenext
	loop
	rsMessages.Movefirst
  end if
end if
%>

<html>
<head>

<script language="JavaScript">
<!-- hide on

function popup(popupfile,winheight,winwidth)
{
open(popupfile,"PopupWindow","resizable=no,height=" + winheight + ",width=" + winwidth + ",scrollbars=no");
}

// hide off -->
</script>

<title>Homepage</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2> 
<h3 align="center">Área de membros do Miragem</h3>
<table width="180" cellpadding=3 cellspacing=0 border=1 bordercolor="#000000" align="right">
  <%'If they're not logged in, then display a login box
if username = "" then%>
  <tr> 
    <th bgcolor="#000000"><font face="arial,helvetica" size="2" color="#FFFFFF">Cadastro 
      e login</font></th>
  </tr>
  <tr> 
    <form name="login" action="signin.asp" method="post">
      <input type="hidden" name="page" value="index.asp">
      <td align="center" bgcolor="#FFFFFF"> <font face="arial,helvetica" size=1><b> 
        Usu&aacute;rio : 
        <input type="text" name="username" size="12" style="font-size: 8pt; font-family: Tahoma, Verdana Arial, Helvetica, sans-serif;">
        <br>
        Senha : 
        <input type="password" name="password" size="12" style="font-size: 8pt; font-family: Arial, Helvetica, sans-serif;">
        <br>
        Lembrar senha: 
        <input type="checkbox" name="stayloggedin" value="yes">
        <br>
        <input type="submit" value="Login" style="font-size: 8pt; font-family: Arial, Helvetica, sans-serif;">
        <br>
        <a href="signupform.asp"><font size="2">Ainda n&atilde;o sou membro<br>
        inscreva-me agora!</font></a> </b></font> </td>
    </form>
  </tr>
  <%'If they are, show a mini profile box plus a sign out link
else%>
  <tr> 
    <th bgcolor="#000000"><font face="arial,helvetica" size="2" color="#FFFFFF">Bem 
      vindo de volta!</font></th>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <font face="arial,helvetica" size=1><b> <img src="icons/<%=rsUser("icon")%>.gif" width=50 height=50 align="right"> 
      <font size="2">Bem vindo <%=username%>.</font><br>
      <a href="javascript:popup('profile.asp',370,275)">Ver/editar perfil</a><br>
      <a href="inbox.asp">Minha caixa de mensagens(<%=newcount%> novas)</a><br>
      <a href="userview.asp">Ver lista de membros</a><br>
      <a href="signout.asp">Sair da &aacute;rea de membros</a> </b></font> </td>
  </tr>
  <%
rsUser.close
set rsUser = nothing
rsMessages.close
set rsMessages = nothing
conn.close
set conn = nothing
end if
%>
</table>
<p><b>Esta &eacute; mais nova se&ccedil;&atilde;o do portal Miragem onde nossos 
  visitantes podem se cadastrar e entrar em contato com outros membros cadastrados.</b></p>
<p>O cadastro &eacute; simples e descomplicado e &eacute; necess&aacute;rio preencher 
  um pequeno formul&aacute;rio clicando no quadro ao lado.</p>
<p>Estando cadastrado voc&ecirc; poder&aacute; entrar em contato enviando mensagens 
  para outros usu&aacute;rios cadastrados de forma r&aacute;pida e segura e poder&aacute; 
  tamb&eacute;m verificar quem est&aacute; cadastrado na &aacute;rea &quot;Lista 
  de membros&quot;..</p>
<p><b><i><font color="#FF0033">Todos os membros cadastrados estar&atilde;o concorrendo 
  mensalmente &agrave; uma an&aacute;lise astrol&oacute;gica completa (hor&oacute;scopo 
  cigano, oriental, ocidental ,anjos) que ser&aacute; enviada pelo e-mail em formato 
  word.</font></i></b></p>
<p>N&atilde;o perca tempo e venha fazer parte da nossa comunidade.</p>
<p>Em caso de d&uacute;vidas entre em contato conosco pelo e-mail <a href="mailto:miragem@esomiragem.com.br"> 
  miragem@esomiragem.com.br</a>. </p>
<p>&nbsp;</p>
</font> 
</body>
</html>
