<%
Option Explicit
Dim sql, rsUser, username, password, firstname, surname, email, dob, sex, notfilled(5), badflag, count, passwordLength, calltype, starsign, dobmonth, dobday, sendemail

if Request.Cookies("philoginadmin") = "True" then

username = Request.Form("username")

'Assign form values to variables

password = Request.Form("password")
firstname = Request.Form("firstname")
surname = Request.Form("surname")
email = Request.Form("email")
sex = Request.Form("sex")
dob = Request.Form("birth_day") & "/" & Request.Form("birth_month") & "/" & Request.Form("birth_year")

if Request.Form("sendemail") = "True" then
	sendemail = True
else
	sendemail = False
end if

'Check everything's been filled in, badflag determines whether error function is called
badflag = 0

if firstname = "" then
	notfilled(0) = "First Name"
	badflag = 1
end if
if surname = "" then
	notfilled(1) = "Surname"
	badflag = 1
end if
if email = "" then
	notfilled(2) = "Email"
	badflag = 1
end if
if sex = "" then
	notfilled(3) = "Sex"
	badflag = 1
end if
if IsDate(dob) = "False" then
	notfilled(4) = "Date of Birth"
	badflag = 1
end if
if password = "" then
	notfilled(5) = "Password"
	badflag = 1
end if

if badflag = 1 then
	signuperror()
end if

'Open connection and insert user details into the database
%>
<!--#include file="conn.asp"-->
<%
getstarsign()

Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.open "users WHERE username = '" & username & "'", conn, 3, 3
rsUser("password") = password
rsUser("firstname") = firstname
rsUser("surname") = surname
rsUser("email") = email
rsUser("dob") = dob
rsUser("starsign") = starsign
rsUser("sex") = sex
rsUser("sendemail") = sendemail
rsUser.Update

rsUser.close
set rsUser = nothing
conn.close
set conn = nothing

Response.Redirect("philogin_admin-profile.asp?username=" & username & "&updated=true")

else
Response.Redirect("philogin_admin.asp")
end if
%>

<%
Function getstarsign()
dobmonth = Request.Form("birth_month")
dobday = Request.Form("birth_day")
Select Case dobmonth
	Case 1
		if dobday < 21 then
			starsign = "Capricorn"
		else
			starsign = "Aquarius"
		end if
	Case 2
		if dobday < 20 then
			starsign = "Aquarius"
		else
			starsign = "Pisces"
		end if
	Case 3
		if dobday < 21 then
			starsign = "Pisces"
		else
			starsign = "Aries"
		end if
	Case 4
		if dobday < 21 then
			starsign = "Aries"
		else
			starsign = "Taurus"
		end if
	Case 5
		if dobday < 22 then
			starsign = "Taurus"
		else
			starsign = "Gemini"
		end if
	Case 6
		if dobday < 22 then
			starsign = "Gemini"
		else
			starsign = "Cancer"
		end if
	Case 7
		if dobday < 24 then
			starsign = "Cancer"
		else
			starsign = "Leo"
		end if
	Case 8
		if dobday < 24 then
			starsign = "Leo"
		else
			starsign = "Virgo"
		end if
	Case 9
		if dobday < 24 then
			starsign = "Virgo"
		else
			starsign = "Libra"
		end if
	Case 10
		if dobday < 24 then
			starsign = "Libra"
		else
			starsign = "Scorpio"
		end if
	Case 11
		if dobday < 23 then
			starsign = "Scorpio"
		else
			starsign = "Sagittarius"
		end if
	Case 12
		if dobday < 22 then
			starsign = "Sagittarius"
		else
			starsign = "Capricorn"
		end if
End Select
End Function
%>

<%Function signuperror()%>
<html>
<head>
<title>Your Profile</title>
</head>
<body bgcolor="#FFFFFF" link="#DD0000" vlink="#DD0000" alink="#000000">
<font face="arial,helvetica" size=2>

<p><b>Voc&ecirc; n&atilde;o preencheu os seguintes dados corretamente:</b></p>

<p>
<%for count = 0 to 5%>
  <%if notfilled(count) <> "" then%>
<b><%=notfilled(count)%></b><br>
  <%end if%>
<%next%>
</p>

<p><a href="javascript:self.history.go(-1)">Tente novamente</a></p>

</font>
</body>
</html>
<%Response.end
End Function%>
