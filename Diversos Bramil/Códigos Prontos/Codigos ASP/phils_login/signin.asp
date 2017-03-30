<%
Option Explicit
Dim sql,rsUser,username,password,page,stayloggedin,LConnectString,lconn,sqlflag

'Grab the submitted variables (page is the page they've come from, set by the hidden variable at the login box)
username = Request.Form("username")
password = Request.Form("password")
page = Request.Form("page")
stayloggedin = Request.Form("stayloggedin")

if page = "" then
	page = "index.asp"
end if

'Check no s**t is trying to hack in using SQL commands
if InStr(username, "'") or InStr(username, """") or InStr(username, "=") or InStr(password, "'") or InStr(password, """") or InStr(password, "=") then
	sqlflag = True
end if

'Open connection
%>
<!--#include file="conn.asp"-->
<%

'Get a recordset corresponding to the submitted username and password
sql = "SELECT username FROM users WHERE username = '" & username & "' AND password = '" & password & "'"
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open sql, conn, 3, 3

'If there was a valid recordset there, then send them back to the page they came from, with the username cookie set
If (not rsUser.BOF) and (not rsUser.EOF) and sqlflag <> True then
  Response.Cookies("username") = rsUser("username")
  'If the user wants to stay logged in all the time, then we'll set the cookie with a far-away expiry date
  if stayloggedin = "yes" then
	Response.Cookies("username").expires = #1/1/2010#
  end if
  rsUser.close
  set rsUser = nothing
  conn.close
  set conn = nothing
  Response.Redirect(page)
end if

'Otherwise, display an invalid entry screen
rsUser.close
set rsUser = nothing
conn.close
set conn = nothing%>

<html>
<head>
<title>Invalid entry</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2 color="#000000">

<h3>Nome de usu&aacute;rio/senha inv&aacute;lidos</h3>

<p><a href="javascript:self.history.go(-1)"><b>Tente novamente</b></a></p>

</font>
</body>
</html>
