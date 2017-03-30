<%
if request.form("reg") = "yes" then
'get variables from form
	usn = request.form("usn")
	pwd = request.form("pwd")
	pwd2 = request.form("pwd2")
	fname = request.form("fname")
	lname = request.form("lname")
	email = request.form("email")

'check varables are present and correct
	if usn = "" then
		err = 1
		errmess = "You didnt insert a username<br>"
	end if
	if pwd = "" then
		err = 1
		errmess = errmess & "You didnt insert a password<br>"
	end if
	if pwd2 = "" then
		err = 1
		errmess = errmess & "You didnt insert a password again<br>"
	end if 
	if fname = "" then
		err = 1
		errmess = errmess & "You didnt insert your first name<br>"
	end if
	if lname = "" then
		err = 1
		errmess = errmess & "You didnt insert your last name<br>"
	end if
	if InStr(email,"@") = 0 or InStr(email,".") = 0 or email = "" then
		err = 1
		errmess = errmess & "You didnt enter a valid email address<br>"
	end if
	if pwd <> pwd2 then
		err = 1
		errmess = errmess & "Your passwords dont match<br>"
	end if

		if err = 0 then
			set dataconn = server.createobject ("ADODB.connection")
			set rs1 = server.createobject ("ADODB.recordset")
			dataconn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\test1\login\db2.mdb"
			MYSQL = "SELECT * FROM members WHERE usern = '" & usn & "'"
			rs1.open MYSQL, dataconn, 1, 3
				if not rs1.EOF or not rs1.BOF then 'username exists already
					errmess = "Your choosen Username already exists"
				else
				
					rs1.AddNew
					rs1.Fields("usern") = usn 
					rs1.Fields("passwo") = pwd
					rs1.Fields("fname") = fname
					rs1.Fields("lname") = lname
					rs1.Fields("email") = email
					rs1.Update
					response.redirect "login.asp"
				end if
			rs1.close()
			dataconn.close()
		end if
end if

%>


<html>
<head>
<title>Registration Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
Please fill in the form below to register
<form name="form1" method="post" action="reg.asp">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="16%" align="right">Username </td>
      <td width="23%"> 
        <input type="text" name="usn" value="<%= usn %>">
      </td>
      <td width="12%" align="right">Password </td>
      <td width="49%"> 
        <input type="password" name="pwd">
      </td>
    </tr>
    <tr> 
      <td width="16%" align="right">Password Again </td>
      <td width="23%"> 
        <input type="password" name="pwd2">
      </td>
      <td width="12%" align="right"> Email </td>
      <td width="49%"> 
        <input type="text" name="email" value="<%= email %>">
      </td>
    </tr>
    <tr> 
      <td width="16%" height="27" align="right">First Name </td>
      <td width="23%" height="27"> 
        <input type="text" name="fname" value="<%= fname %>">
      </td>
      <td width="12%" height="27" align="right">Last Name </td>
      <td width="49%" height="27"> 
        <input type="text" name="lname" value="<%= lname %>">
      </td>
    </tr>
    <tr>
      <td width="16%" height="27" align="right">&nbsp;</td>
      <td width="23%" height="27">
        <input type="submit" name="Submit" value="Submit">
        <input type="reset" name="Submit2" value="Reset">
        <input type="hidden" name="reg" value="yes">
      </td>
      <td width="12%" height="27" align="right">&nbsp;</td>
      <td width="49%" height="27">&nbsp;</td>
    </tr>
  </table>
</form>
<%= errmess %>
</body>
</html>
