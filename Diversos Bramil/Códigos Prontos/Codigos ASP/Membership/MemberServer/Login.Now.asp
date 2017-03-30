<%
Dim strSQL				' Structured Query Language
Dim objConn				' Database Connection
Dim objRs				' Recordset
Dim strLoginName		' Name user logs on with
Dim strLoginPassword	' Password to login with
Dim lngMemberID			' MemberID assigned to user account

' Grab Form Data
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

' Open Database
Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")
objConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("Members.mdb")

' Look for User
strSQL = "SELECT MemberID FROM Members WHERE LoginName = '" & strLoginName & "' AND LoginPassword = '" & strLoginPassword & "'"
Set objRs = objConn.Execute(strSQL)

' Notify visitor if account was found.
If objRs.EOF Then
	Response.Write("Login Failed")
Else
	lngMemberID = objRs(0)
	Response.Write("Login Succeeded, MemberID = " & lngMemberID)
End If

' Garbage Collection
Set objRs = Nothing
objConn.Close
Set objConn = Nothing
%>