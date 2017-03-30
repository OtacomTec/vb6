<%
Dim strSQL				' Structured Query Language
Dim objConn				' Database Connection
Dim strLoginName		' Name user logs on with
Dim strLoginPassword	' Password to login with

' Grab form data
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

' Open Database Connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("Members.mdb")

' Add user
strSQL = "INSERT INTO Members(LoginName, LoginPassword) VALUES('" & strLoginName & "', '" & strLoginPassword & "')"
Call objConn.Execute(strSQL)

' Garbage Collection
objConn.Close
Set objConn = Nothing

' Notify visitor of success
Response.Write("Registration complete")
%>
