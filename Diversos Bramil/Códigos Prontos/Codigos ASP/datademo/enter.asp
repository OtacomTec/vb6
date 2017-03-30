
<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Buffer = True%>
<%
Dim name
Dim pass



name  = CStr(Request.QueryString("id"))
pass = CStr(Request.QueryString("value"))

%><font face="verdana" size="2">
You have succesfully logged in. So your cookie is set<br>
<font face="courier new" size="2"><br><br>
username:  <%=name%><br>
password:   <%=pass%>

<%

'Response.Buffer = True
'sd 
%>
