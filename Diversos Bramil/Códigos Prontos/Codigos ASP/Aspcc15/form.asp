<!-- Begin removable test code -->
<h2 align="center">ASP Client Check 1.4</h2>
<div align="center">
Test Users:
<table cellspacing="2" cellpadding="2" border="1">
	<tr><td><b>USERNAME</b></td><td><b>PASSWORD</b></td></tr>
	<tr><td>Jim</td><td>Miles</td></tr>
	<tr><td>Arthur</td><td>Remmy</td></tr>
	<tr><td>Joe</td><td>Horn</td></tr>
	<tr><td>Fred</td> <td>Smith</td></tr>
</table>
<br>
<!-- End removable test code -->

<b>Please enter your Username and Password.</b> <i>(case sensitive)</i>
<BR>
<form action="aspccConfirm.asp" method="post">

	<% If Request.Cookies("Preferences")("Username") = "" Then 'No cookies form %>
	
		Username: <input name=Username type=Text><p>
		Password: <input name=Password type=password><p>
		<input type="checkbox" name="SaveCookie" value="1">
	
	<% Else	'Cookies form %>
	
		Username: <input name=Username value="<%=Request.Cookies("Preferences")("Username")%>" type=Text><p>
		Password: <input name=Password Value="<%=Request.Cookies("Preferences")("Password")%>" type=password><p>
		<input type="checkbox" name="SaveCookie" value="1" checked>
		
	<% End If %>

Save Username and Password for future visits?<p>
<Input type=submit value=Submit>
</form>

