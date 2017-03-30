<% 
	' ********************************************
	' Easy Password protection for your pages - 
	' see the readme.txt in this zip file for
	' more details on how to implement the script
	' This version copyright dotdragnet.co.uk
	' ********************************************


	' Set your password here
	' Make sure it has quote marks around it	
	pass = "password-goes-in-here"


	' ********************************************
	' DO NOT MODIFY BELOW THIS LINE
	' ********************************************
	URL = Request.ServerVariables("SCRIPT_NAME")
	if request.form("button") = "logout" then
		session("password")=""
		session("logon")=""
	end if
	if session("password") = pass then 
		' session set ok, load page
		session("logon")="true"
		displaypage
	else
		if request.form("password") = "" then
			' form not posted and session not set, display logon box
			displaylogon
		else
			' something posted, check it
			session("password") = request.form("password")
			if session("password") = pass then 
				' session set ok, load page
				session("logon")="true"
				displaypage
				
			else
				displaylogon
			end if
		end if
	end if
%>
<% function displaylogon()%>
<body bgcolor="#cccccc">
<div align="center">
  <center>
  <table border="0" height=100% width=100%>
    <tr>
      <td valign=center>
       <form method="POST" action="<%=URL%>" >
                 <p align="center"><b>Please enter the password to enter this page : </b><input type="password" name="password" size="20"><input type="submit" value="Login" name="button"></p>
        </form>
        <p align="center"> </td>
    </tr>
  </table>
  </center>
</div>
<%end function%>



'	Having trouble with this script? Why not check out the dotdragnet web builder forum 
'	Where you can get more advice and help - http://www.dotdragnet.co.uk/forum



