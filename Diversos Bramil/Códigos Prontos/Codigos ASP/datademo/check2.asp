<%
USERNAME = Replace(Request("USERNAME"),"'","")
wachtw = Replace(Request("pass"),"'","")
status = Request("status")


'login = "<form method=""POST"" action="" & THISPAGE & ""><input type=""hidden"" name=""status"" value=""check""><br><br><input type=""text"" name=""USERNAME""><br><input type=""text"" name=""pass""><input type=""submit"" value=""Login""></form>"


dim OKEE 
OKEE = 0
dim item
Dim strSQL
Dim objConn
Dim objRec2
Dim StrConnect	
THISPAGE = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")

	strConnect = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("db/data.mdb") &";DefaultDir="& Server.MapPath(".") &";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5"	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set CmdCheckUser = Server.CreateObject("ADODB.Recordset")
	objConn.Open strConnect
	SQL = "SELECT Naam, ID, pass FROM data WHERE (Naam = '" & USERNAME & "')"
	CmdCheckUser.Open SQL, objConn
	
		
if status = "check" then


	If CmdCheckUser.EOF And CmdCheckUser.BOF Then
		OKEE = 0 %> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">no valid username!</font> 
<form method="POST" action="<%=THISPAGE%>"><input type="hidden" name="status" value="check">
  <table width="192" border="0" cellspacing="1" cellpadding="0" height="88" bgcolor="#000000">
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="text" value="" name="USERNAME">
      </td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="text" name="pass">
      </td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="submit" value="Login" name="submit">
      </td>
    </tr>
  </table>
  </form>
		
<%
	else


		dim pass
		pass = CmdCheckUser("pass")


		if not Request("pass") <> CmdCheckUser("pass") then



			For Each Item in Request.Cookies
			Response.Cookies(Item).Expires = Date() - 1
			Next

			Response.Cookies(USERNAME) = pass
			Response.Cookies(USERNAME).Expires = Date() + 100


					x =0
	For Each Item in Request.Cookies
	x = x + 1	%> <% pas = Request.Cookies(Item) %> <% na = Item %> 


<table width="330" border="0" cellspacing="1" cellpadding="0" height="20" bgcolor="#000000">
  <tr bgcolor="#FFCC00"> 
    <td><center><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="enter.asp?id=<%=na%>&value=<%=pas%>"><%=na%></a></font></center></td>
    <td><center><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="delcook.asp?id=<%=na%>">delete cookie!</a> </font></center></td>
  </tr>
</table>






<%
	Next
			OKEE = 1
		else
			%><font face="Verdana, Arial, Helvetica, sans-serif" size="2">password 
for: <%=CmdCheckUser("naam")%></font> 
<form method="POST" action="<%=THISPAGE%>"><input type="hidden" name="status" value="check">
  <br>
  <input type="hidden" value="<%=USERNAME%>" name="USERNAME">
  <table width="192" border="0" cellspacing="1" cellpadding="0" height="61" bgcolor="#000000">
    <tr bgcolor="#FFCC00" align="center"> 
      <td height="21"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=USERNAME%></font></td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td height="33"> 
        <input type="text" name="pass">
      </td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td height="41"> 
        <input type="submit" value="Login" name="submit">
      </td>
    </tr>
  </table>
  <br>
</form> 
<%
		end if 
	end if


else


	x =0
	For Each Item in Request.Cookies
	x = x + 1	%> <% pas = Request.Cookies(Item) %> <% na = Item %> 


<table width="330" border="0" cellspacing="1" cellpadding="0" height="20" bgcolor="#000000">
  <tr bgcolor="#FFCC00"> 
    <td><center><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="enter.asp?id=<%=na%>&value=<%=pas%>"><%=na%></a></font></center></td>
    <td><center><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="delcook.asp?id=<%=na%>">verwijder cookie!</a> </font></center></td>
  </tr>
</table>






<%
	Next
	
	if x = 0 then
		%><font face="Verdana, Arial, Helvetica, sans-serif" size="2">no 
cookies </font>
<form method="POST" action="<%=THISPAGE%>"><input type="hidden" name="status" value="check">
  <table width="192" border="0" cellspacing="1" cellpadding="0" height="88" bgcolor="#000000">
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="text" value="" name="USERNAME">
      </td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="text" name="pass">
      </td>
    </tr>
    <tr bgcolor="#FFCC00" align="center"> 
      <td> 
        <input type="submit" value="Login" name="submit">
      </td>
    </tr>
  </table>
</form>
		<%

	end if


end if

%>

