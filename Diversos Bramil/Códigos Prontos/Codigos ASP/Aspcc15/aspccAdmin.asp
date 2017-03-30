<% Option Explicit 'forces variables to be defined prior to use%>
<% Response.Buffer = True 'waits till page is processed before sending response%>
<% On Error Resume Next' proceeds with script if error found (error subroutine catches errors) %>
<!--#include file="adovbs.inc"-->
<!--#include file="aspccConfig.asp"-->
<%

'Check that Admin status has been set (by aspccConfirm.asp)
If Session("Admin") <> 1 Then Response.End

Dim objConn, objRS
Dim strSQL

'Set connection object
Set objConn = Server.CreateObject("ADODB.Connection")

'Open connection object strDBConn (declared and defined in config file)
objConn.open strDBConn

'Check for errors in objConn
subErrorCheck

'Set Recordset
Set objRS = Server.CreateObject("ADODB.Recordset")


'==================================================================================================
' Root Processes
'==================================================================================================
If Request.QueryString("Delete") <> "" Then

	'Execute Delete subroutine.
	subDeleteRecord

ElseIf Request.QueryString("Edit") <> "" Then

	'Execute Edit subroutine
	subEditRecord
	
ElseIf Request.QueryString("Add") <> "" Then

	'Execute Add subroutine
	subAddRecord

Else
	
	'Execute List subroutine
	subListUserInfo

End If





objRs.Close
Set objRs = Nothing 
objConn.Close
Set objConn = Nothing



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'SUBROUTINES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'==================================================================================================
' Subroutine - Delete Record
'==================================================================================================
Sub subDeleteRecord

Dim i, intFieldCount

	'Define SQL Query
	strSQL = "SELECT * FROM [users] WHERE [ID]=" & Request.Querystring("Delete")
	
	'Open recordset passing the SQL to the connection object.
	'Open as Static to be able to execute more Move commands in the recordset.
	objRS.open strSQL, objConn, adOpenStatic, adLockOptimistic
	
	'Check for errors in objConn
	subErrorCheck
	
	intFieldCount = objRS.Fields.Count
	
	If Request.Querystring("Confirmed") = 1 Then
	'Delete record:
		
		objRS.Delete
		
		'Check for errors in objConn
		subErrorCheck
		
		%>
		<p align="center">Record #<%= Request.Querystring("Delete") %> has been deleted.</p>
		<p align="center"><a href="aspccAdmin.asp">Return to Admin</a></p>
		<%
	
	Else
	'Confirm delete prompt:
		
		%>
		<p align="center">Are you sure you wish to delete the following account?</p>
		<table align="center" cellspacing="0" cellpadding="3" border="1">
		<tr>
		<% For i = 0 To intFieldCount -1 %>
			<th><%= objRS(i).Name %></th>
		<% Next %>
		</tr>
		<tr>
		<% For i = 0 To intFieldCount -1 %>
			<td><%= objRS(i)%></td>
		<% Next %>	
		</tr>
		</table>
		<p align="center">
		<a href="aspccAdmin.asp?Delete=<%= Request.Querystring("Delete") %>&Confirmed=1">Yes</a> |
		<a href="aspccAdmin.asp">No</a>
		</p>
		<%
		
	End If

End Sub


'==================================================================================================
' Subroutine - Edit Record
'==================================================================================================
Sub subEditRecord

Dim i, intFieldCount

	'Define SQL Query
	strSQL = "SELECT * FROM [users] WHERE [ID]=" & Request.Querystring("Edit")
	
	'Open recordset passing the SQL to the connection object.
	'Open as Static to be able to execute more Move commands in the recordset.
	objRS.open strSQL, objConn, adOpenStatic, adLockOptimistic
	
	'Check for errors in objConn
	subErrorCheck
	
	intFieldCount = objRS.Fields.Count
	
	If Request.Form("Update") = 1 Then
		
		If Request.Form("Username") = "" or Request.Form("Password") = "" Then
		
		%>
		<p align="Center">
		<font color="#FF0000">Error.</font><br><br> Username and Password fields cannot be blank.<br><br>
		<a href='#1' onClick='history.back()'>Back</a>
		</p>
		<%
		
		Else
		
		'Update the database
		objRS("Username") = Trim(Request.Form("Username"))
		objRS("Password") = Trim(Request.Form("Password"))
	
			If Trim(Request.Form("Email")) = "" Then
				objRS("Email") = "n/a"
			Else
				objRS("Email") = Trim(Request.Form("Email"))
			End If
		
		objRS.Update
		
		'Check for errors in objConn
		subErrorCheck
		
		%>
		<p align="Center">
		Your update has been processed succesfully.<br><br>		
		<a href="aspccAdmin.asp">Return to Admin</a>
		</p>
		<%
		
		End If
	Else
	
		%>
		<form method="post" action="aspccAdmin.asp?Edit=<%= Request.Querystring("Edit") %>">
		<table align="center" cellspacing="0" cellpadding="3" border="0">
		<tr><td align="center" colspan="2"><strong>Please edit the information below:</strong></td></tr>
		<% For i = 1 To intFieldCount -1 %>
		<tr>
			<td><%= objRS(i).Name %></td>
			<td><Input type="text" name="<%= objRS(i).Name %>" value="<%= objRS(i) %>"></td>		
		</tr>
		<% Next %>
		<tr>
			<td colspan="2" align="center">
			<input type="hidden" name="Update" value="1">
			<input type="submit" value="submit">
			</td>
		</tr>
		</table>
		</form>
		<%
	
	End If
	
End Sub


'==================================================================================================
' Subroutine - Add Record
'==================================================================================================
Sub subAddRecord

Dim i, intFieldCount
	
	'Open recordset passing the SQL to the connection object.
	'Open as Static to be able to execute more Move commands in the recordset.
	objRS.open "users", objConn, adOpenStatic, adLockOptimistic
	
	'Check for errors in objConn
	subErrorCheck
		
	If Request.Form("Add") = 1 Then
	
		If Request.Form("Username") = "" or Request.Form("Password") = "" Then
		
		%>
		<p align="Center">
		<font color="#FF0000">Error.</font><br><br> Username and Password fields cannot be blank.<br><br>
		<a href='#1' onClick='history.back()'>Back</a>
		</p>
		<%
		
		Else
		
		'Update the database
		objRS.AddNew
		objRS("Username") = Trim(Request.Form("Username"))
		objRS("Password") = Trim(Request.Form("Password"))
	
			If Trim(Request.Form("Email")) = "" Then
				objRS("Email") = "n/a"
			Else
				objRS("Email") = Trim(Request.Form("Email"))
			End If
		
		objRS.Update
		
		
		
		'Check for errors in objConn
		subErrorCheck
		
		%>
		<p align="Center">
		The new user had been added succesfully.<br><br>		
		<a href="aspccAdmin.asp">Return to Admin</a>
		</p>
		<%
		End If
		
	Else
	
	intFieldCount = objRS.Fields.Count
	
		%>
		<form method="post" action="aspccAdmin.asp?Add=1<%= Request.Querystring("Edit") %>">
		<table align="center" cellspacing="0" cellpadding="3" border="0">
		<tr><td align="center" colspan="2"><strong>Please enter the information below:</strong></td></tr>
		<% For i = 1 To intFieldCount -1 %>
		<tr>
			<td><%= objRS(i).Name %></td>
			<td><Input type="text" name="<%= objRS(i).Name %>"></td>		
		</tr>
		<% Next %>
		<tr>
			<td colspan="2" align="center">
			<input type="hidden" name="Add" value="1">
			<input type="submit" value="submit">
			</td>
		</tr>
		</table>
		</form>
		<%
		
	End If
		
End Sub


'==================================================================================================
' Subroutine - List User Info
'==================================================================================================
Sub subListUserInfo
	
	Dim i, intFieldCount
	
	'Define SQL Query
	strSQL = "SELECT * FROM users ORDER BY Username"
	
	'Open recordset passing the SQL to the connection object.
	'Open as Static to be able to execute more Move commands in the recordset.
	'
	objRS.open strSQL, objConn, adOpenStatic, adLockOptimistic
	
	'Check for errors in objConn
	subErrorCheck
	
	intFieldCount = objRS.Fields.Count
	
	%>
	<p align="center"><a href="aspccAdmin.asp?Add=1">Add New</a></p>
	<table align="center" cellspacing="0" cellpadding="3" border="1">
	<tr>
	<% For i = 0 To intFieldCount -1 %>
		<th><%= objRS(i).Name %></th>
	<% Next %>
		<th>&nbsp;</th>
	</tr>
	
	<% Do Until objRS.EOF %>
	<tr>
	<% For i = 0 To intFieldCount -1 %>
		<td><%= objRS(i)%></td>
	<% Next %>
		<td>
		<a href="aspccAdmin.asp?Edit=<%= objRS("ID") %>">Edit</a> |
		<a href="aspccAdmin.asp?Delete=<%= objRS("ID") %>">Delete</a>
		</td>
		
	</tr>
	<% objRS.Movenext %>
	<% Loop %>
	
	</table>
	<%
	
End Sub


'===========================================================================================
'Error Display Subroutine
'===========================================================================================
Sub subErrorCheck

If objConn.Errors.Count <> 0 Then

	Dim e
		
		Response.Write "<Div align=center><font color=#FF0000><h2>VBScript Error Ecountered:</h2></font>"
		Response.Write "<table width=640 border=1 cellpadding=2 cellspacing=2>"
		
		If objConn.Errors.Count = 1 Then
			
			Response.Write "<tr><td>ERROR#</td><td>" & Err.Number & "</td></tr>"
			Response.Write "<tr><td>DESCRIPTION</td><td>" & Err.Description & "</td></tr>"
			Response.Write "<tr><td>HELP CONTEXT</td><td>" & Err.HelpContext & "</td></tr>"
			'Response.Write "<tr><td>Help File</td><td>" & Err.HelpFile & "</td></tr>"
			
			If strSQL = "" Then		
			Else
				Response.Write "<tr><td>SQL: </td><td>" & strSQL & "</td></tr>"
			End If
	
		Else
		
			For e = 0 To objConn.Errors.Count -1
				
				If objConn.Errors(e) < 0 Then
				
					Response.Write "Error # "
				
				Else
				
					Response.Write "Warning # "
				
				End If
				
		Response.Write objConn.Errors(e).Number & " - "
		Response.Write objConn.Errors(e).Description & "<br>"
			
			Next
		
		End If
		
		Response.Write "</table>"
		Response.Write "<br><br><a href='#1' onClick='history.back()'>Back</a>"
		Response.Write "</DIV>"
		
		'======================================================================
		'Log Error As Text File
		'======================================================================
		
		Dim objFileSys, objLogFile
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		'Set object to OpenTextFile. Location; (8) = append; (true) = create file if does not exist.
		Set objLogFile = objFileSys.OpenTextFile(Session("sesErrLogURL") & Err.Number & ".txt", 8, True)
		
		'Write to file then close.
		objLogFile.WriteLine "WHEN: " & Date() & " - " & Time()
		objLogFIle.WriteLine "ERROR#: " & Err.Number
		objLogFIle.WriteLine "DESCRIPTION: " & Err.Description
		
			If strSQL = "" Then		
			Else
				objLogFile.WriteLine "SQL: " & strSQL
			End If
			
		objLogFile.WriteLine "IP: " & Request.ServerVariables("REMOTE_HOST")		
		objLogFile.WriteLine "BROWSER: " & Request.ServerVariables("HTTP_USER_AGENT")
		objLogFile.WriteBlankLines(1)
		objLogFile.Close
		
		Response.End
		
	End If
	
End Sub

%>