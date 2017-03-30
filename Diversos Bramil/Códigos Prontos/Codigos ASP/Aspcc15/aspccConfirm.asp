<% Option Explicit 'forces variables to be defined prior to use%>
<% Response.Buffer = True 'waits till page is processed before sending response%>
<!--#include file="aspccConfig.asp"-->
<%

Dim strSQL
Dim objConn, objRS

'End script if no Username posted.
If Request.Form("Username") = "" Then Response.End

'=====================================================================================================
'Check If Admin Login
'=====================================================================================================

If Request.form("Username") = strAdminName Then subAdminCheck

'=====================================================================================================
'Open and search recordset
'=====================================================================================================

'Set "objConn" to database connection object.
Set objConn = Server.CreateObject("ADODB.Connection")

'Open database
objConn.Open strDBConn

'Define SQL variable
strSQL = "SELECT Password FROM users WHERE Username='" & Request.form("Username") & "'"

'Set Recordset
Set objRS = objConn.Execute(strSQL)


'=====================================================================================================
'Check form input results against recordset
'=====================================================================================================

'If recordset contains no records, user does not exist.
If objRS.EOF Then

	'Logging subroutine
	If intLog = 1 Then subLogWrite 0

	Response.Write "User <b>" & Request.form("Username") & "</b> does not exist."

'Check if password is correct
ElseIf objRS("Password") = Request.form("Password") Then
	
	'Logging subroutine
	If intLog = 1 Then subLogWrite 1
	
	'Cookie subroutine
	subWriteCookie
	
	'Custom Code subroutine
	subCustomCode
		
Else

	'Logging subroutine
	If intLog = 1 Then subLogWrite 0
	
	Response.Write "Incorrect password."	

End If

objRs.Close
Set objRs = Nothing 
objConn.Close
Set objConn = Nothing





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' SUBROUTINES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'==================================================================================================
'Admin Login Check
'==================================================================================================
Sub subAdminCheck

	If Request.form("Password") = strAdminPass Then
		
		'Logging subroutine
		If intLog = 1 Then subLogWrite 1
		
		'Declare admin session variable:
		Session("Admin") = 1
		
		'Redirect admin to admin script:
		Response.Redirect("aspccAdmin.asp")
	
	Else
		
		'Logging subroutine
		If intLog = 1 Then subLogWrite 0	
		
		Response.Write "Incorrect password."
		Response.End
	
	End If
	
End Sub


'==================================================================================================
'Cookie subroutine
'==================================================================================================
Sub subWriteCookie

	If Request.Form("SaveCookie") = 1 Then
	
		Response.Cookies("Preferences")("Username") = Request.form("Username")
		Response.Cookies("Preferences")("Password") = Request.form("Password")
		Response.Cookies("Preferences").Expires = "January 1, 2002"
		Response.Write("Cookie saved.<p>")
		
	Else
	
		Response.Cookies("Preferences")("Username") = ""
		Response.Cookies("Preferences")("Password") = ""
	
	End If

End Sub


'==================================================================================================
'Logging Subroutine
'==================================================================================================
Sub subLogWrite(intResult)

	Dim objFileSys, objLogFile
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	'Set object to OpenTextFile. Location, (8) = append, (true) = create file if does not exist.
	Set objLogFile = objFileSys.OpenTextFile(strLogURL, 8, True)
	
	'Write to file then close.
	objLogFIle.Write(intResult & " | ") 'request result of login attempt 
	objLogFile.Write(Request.form("Username") & " | ") 'request name
	objLogFile.Write Request.ServerVariables("REMOTE_HOST") 'request IP info
	objLogFile.Write(" | " & Date() & " | " & Time() & " | ") ' date and time
	objLogFile.WriteLine Request.ServerVariables("HTTP_USER_AGENT") 'request browser info
	objLogFile.WriteBlankLines(1)
	objLogFile.Close	

End Sub
%>