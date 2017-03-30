Easy Password protection

Add the following code to the html page you wish to password protect:

1) between the </head> and <body> tags:

<!-- #include File="passwordtop.asp" --> 
<% function displaypage()%>


2) at the bottom of the html page between the </body> and </html> tags:  

<%end function%>
<!-- #include File="passwordbottom.asp" --> 

Change the extension of your web page to .asp (if it was .html or .htm for example)
Make sure the passwordtop.asp and passwordbottom.asp files are in the same folder

That's it!  Easy!


Having problems with this script ?  Check out the dotdragnet Web Builder forum.

http://www.dotdragnet.co.uk