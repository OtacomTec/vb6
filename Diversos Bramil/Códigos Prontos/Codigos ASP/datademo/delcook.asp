
<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Buffer = True%>
<%

ID = CStr(Request.QueryString("ID"))
Response.Cookies(ID).Expires = Date() - 1
Response.Redirect("check2.asp")

%>
