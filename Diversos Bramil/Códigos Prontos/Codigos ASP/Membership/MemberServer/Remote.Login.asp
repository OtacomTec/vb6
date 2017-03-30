<%
Dim strSQL				' Stuctured Query Language
Dim objConn				' Database Connection
Dim objRs				' Recordset
Dim strLoginName		' Name user logs on with
Dim strLoginPassword	' Password to login with
Dim lngMemberID			' MemberID assigned to account

' Ensure page results do not remain in cache
Response.Expires = 0
Response.ExpiresAbsolute = Now()

' Grab querystring data
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

' Open database connection
Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")
objConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("Members.mdb")

' Search for Member Account
strSQL = "SELECT MemberID FROM Members WHERE LoginName = '" & strLoginName & "' AND LoginPassword = '" & strLoginPassword & "'"
Set objRs = objConn.Execute(strSQL)

' Notify viewer/server of results
If objRs.EOF Then
	Response.Write("Login Failed")
Else
	lngMemberID = objRs(0)
	Response.Write(lngMemberID)
	
	' HACK: Errors occur if result is a single number
	Response.Write(".")
	
End If

'Response.Write "<HR>QueryString: " & Request.QueryString & "<HR>"
'Response.Write "<HR>Form: " & Request.Form & "<HR>"

' Garbage Collection
Set objRs = Nothing
objConn.Close
Set objConn = Nothing

Response.End
%>