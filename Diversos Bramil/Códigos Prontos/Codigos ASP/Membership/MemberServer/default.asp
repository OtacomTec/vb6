<H1>Membership Services Demonstration</H1>
<H2>version 1.0</H2>
<P>
	Below are links to the basic functionality of any
	membership/authentication services offered by most
	online providers.  Please take a look at the code
	for a complete understanding of the inner workings
	of this idea.
</P>
<P>
	If these files are not within http://localhost/MemberServer/
	directory, then you will need to edit /MemberServer/Remote/Login.Now.asp
	file to reflect the actual location of /MemberServer/Remote.Login.asp
</P>
<P>
	This test is initially setup for demonstrating the centralized and remote
	server on the same machine.  For better proof of concept, it is suggested
	to move the files in the /MemberServer/Remote/ directory to a seperate
	web server on the internet and modify it to point to the /MemberServer/Remote.Login.asp
	page.
</P>
<UL>
	<LI><A href="Register.asp">Register</A></LI>
	<LI><A href="Login.asp">Login</A></LI>
</UL>
<UL>
	<LI><A HREF="Remote/Login.asp">Remote Login</A></LI>
</UL>