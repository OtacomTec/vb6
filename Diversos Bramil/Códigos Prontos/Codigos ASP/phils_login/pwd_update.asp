<%
Option Explicit
Dim sql, rsUser, username, oldpassword, newpassword, newpasswordconfirm, passwordlength, notfilled(4), calltype, icon

username = Request.Cookies("username")

'Assign form values to variables

oldpassword = Request.Form("oldpassword")
newpassword = Request.Form("newpassword")
newpasswordconfirm = Request.Form("newpasswordconfirm")
icon = Request.Form("icon")

'Ensure new password have been filled in
if newpassword = "" then
	errorfunction("nonew")
end if

'Check password length is between 5 and 15 characters long
passwordLength = Len(newpassword)
if passwordLength < 5 or passwordLength > 15 then
	errorfunction("length")
end if

'Check password and confirmed password are the same
if newpassword <> newpasswordconfirm then
	errorfunction("confirm")
end if

'Open connection and insert user details into the database
%>
<!--#include file="conn.asp"-->
<%
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.open "users WHERE username = '" & username & "'", conn, 3, 3

if rsUser("password") <> oldpassword then
	errorfunction("wrongpwd")
else
	rsUser("password") = newpassword
	rsUser.Update
	
	rsUser.close
	set rsUser = nothing
	conn.close
	set conn = nothing
	
	Response.Redirect("changepwd.asp?updated=true&icon=" & icon)
end if
%>

<%Function errorfunction(calltype)%>
<html>
<head>
<title>Your Profile</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2>

<%
if calltype = "nonew" then
	Response.Write("<p>You have not entered a new password.</p>")
elseif calltype = "length" then
	Response.Write("<p>Your password must be between 5 and 15 characters long.</p>")
elseif calltype = "confirm" then
	Response.Write("<p>Your new password and confirmed password are not the same.</p>")
elseif calltype = "wrongpwd" then
	Response.Write("<p>You have entered an incorrect old password.</p>")
	rsUser.close
	set rsUser = nothing
	conn.close
	set conn = nothing
end if
%>

<p><a href="javascript:self.history.go(-1)">Tente novamente</a></p>

</font>
</body>
</html>
<%Response.end
End Function%>
