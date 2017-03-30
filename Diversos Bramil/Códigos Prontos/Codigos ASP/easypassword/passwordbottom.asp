<% 
	if session("logon") = "true" then 
		' session set ok, load page
		displaylogout
	end if
%>
<% function displaylogout()%>
<br>
<div align="center">
  <center>
  <table border="0" width=100%>
    <tr>
      <td valign=center>
        <form method="POST" >
                 <p align="center"><input type="submit" value="logout" name="button"></p>
        </form>
        <p align="center"> </td>
    </tr>
  </table>
  </center>
</div>
<%end function%>


