<%
if Request.Cookies("username") = "" then
  Response.Redirect("nologin.asp?page=membersonly.asp")
end if
%>
<html>
<head>
<title>Members only page</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2>

<h2>&Aacute;rea somente para membros</h2>

<p>Se voc&ecirc; ainda n&atilde;o se inscreveu, voc&ecirc; ser&aacute; direcionado 
  para a &aacute;rea de inscri&ccedil;&atilde;o.</p>

<p><a href="index.asp">Voltar &agrave; p&aacute;gina inicial</a></p>

</font>
</body>
</html>