<% response.buffer = true %>
<html>
<head>
<title>Logon Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" >
<p> 
  
</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="19" align="center">
  <tr> 
    <td height="54" align="left"> 
<%
' BEGIN LOGON PROCEDURE
dologin = request.form("login")
usn = request.form("username")
psw = request.form("password")
if request.cookies("logoncookie")("cookname")<>"" then
	usn = request.cookies("logoncookie")("cookname")
	psw = request.cookies("logoncookie")("cookpass")
	dologin = "login"
end if
if request.form("logoff")="logoff" then
	dologin = "no"
	session("logon")="no"
	session("usn") = ""
	session("admin") = "no"
	response.cookies("logoncookie").expires = date  -1
end if
	'decide whether to login or not
	if dologin ="login" then
		set dataconn = server.createobject ("ADODB.connection")
		set rs1 = server.createobject ("ADODB.recordset")
		dataconn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\test1\login\db2.mdb"
		MYSQL = "SELECT * FROM members WHERE usern = '" & usn & "'"
		rs1.open MYSQL, dataconn, 1, 3
		if not rs1.EOF or not rs1.BOF then 'username checker
			if psw = (rs1("passwo")) then  'check password
				session("logon") = "yes"
				session("usn") = usn
				if rs1("level") = "admin" then session("admin") = "yes"
				if request.form("rem") = "yes" then 'checkbox and create cookie
					response.cookies("logoncookie").Expires = date + 31
					response.cookies("logoncookie")("cookname")= usn
					response.cookies("logoncookie")("cookpass")= psw
				end if
			else
				session("logon") = "no" 'incorect password error
				errmess="password incorrect"
			end if
		else
			errmess="Incorect Username" 'incorrect username error
		end if
		rs1.close()
		dataconn.close()
		
	end if
'show user logged in
if session("logon") = "yes" then
call logonyes
else
'show login required
call logonno
response.write errmess
end if
' END LOGIN PROCEDURE
%>
    </td>
  </tr>
</table>
<% function logonno() %>
<form name="form1" method="post" action="login.asp">
              Username 
              <input type="text" name="username" size="17">
              <br>
              Password 
              <input type="password" name="password" size="17">
              <br>
              Remember me 
              <input type="checkbox" name="rem" value="yes">
              <input type="submit" name="login" value="login">
</form>
<p>Please click <a href="reg.asp">here</a> to register</p>
<% end function 

function logonyes() 
response.write "You are logged on as " & session("usn")
if session("admin") = "yes" then response.write "<br>Admin Level Logon"
%>
<form name = "form1" method="post" action="login.asp">
              <input type="submit" name="logoff" value="logoff">
</form>
<% end function %>
</body>
</html>
