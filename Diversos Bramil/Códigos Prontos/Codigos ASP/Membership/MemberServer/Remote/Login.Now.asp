<%
Dim strSQL				' Structured Query Language
Dim strLoginName		' Name user logs on with
Dim strLoginPassword	' Password to login with
Dim lngMemberID			' MemberID assigned to user account
Dim objSpider			' Spider to request HTML page that logs user into system
Dim strResponseText		' Response returned from remote server
Dim strQueryString		' QueryString

' URL to page that allows other servers to log users against centralized
' database.  (This may need to be changed)
Const strURL = "http://localhost/MemberServer/Remote.Login.asp"

' Grab posted form data
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

strQueryString = _
	"LoginName=" & Server.URLEncode(strLoginName) & _
	"&LoginPassword=" & Server.URLEncode(strLoginPassword)
	
' Log user into remote location and get results
Set objSpider = Server.CreateObject("Microsoft.XMLHTTP")
' other versions ...
' MSXML2.XMLHTTP.3.0, MSXML2.ServerXMLHTTP, Microsoft.XMLHTTP

With objSpider
	Call .Open("POST", strURL, False, "", "")
	Call .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	Call .Send(strQueryString)
	strResponse = .ResponseText
End With
Set objSpider = Nothing

' Notify user of success/failure
If IsNumeric(strResponse) And Not strResponse = "" Then
	lngMemberID = strResponse
	Response.Write("MemberID = " & lngMemberID)
Else
	Response.Write("Login Failed: " & strResponse)
End If
%>