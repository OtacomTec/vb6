<%
'ASP Client Check - Configuration File

Dim strDBConn, intLog, strLogURL, strAdminName, strAdminPass

'Database connection. Edit the url of the following line to point to the "clients.mdb" file on
'the server computer, or replace everything between the double quotes with the name of your system DSN.
strDBConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\www\scripts\aspcc15\clients.mdb;"

'Set intLog to "1" if you want to keep a log of attempts to enter. 
intLog = 1
	
'Location of log file. For added security, you can make this a location inaccessible via the web.
strLogURL = "e:\www\scripts\aspcc15\cclog.txt"

'Administration username and password. Set these variables only if using the admin script.
'Administration script requires cookies to be enabled in browser for session variable.
strAdminName = ""
strAdminPass = ""


'======================================================================================================
'Your custom subroutine
'======================================================================================================
'Please test the script with this default subroutine prior to modifying it. 
'The code contained in this subroutine will execute when your clients login successfully.
'Alter only the code within the "Sub" and "End Sub" delimiters.
Sub subCustomCode

	Response.Write "<p><b>" & Request("Username") & "</b>, please enjoy your stay!<p>"
	Response.Write "<i>Your custom function is executed at this point.</i>"
	
End Sub
%>