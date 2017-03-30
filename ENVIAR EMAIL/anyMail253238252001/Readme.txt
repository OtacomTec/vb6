anyMail v1.1.5
==================

By       : Saurabh Gupta
	   -------------
E-mail	 : saurabh_gupta@india.com
	   -----------------------	
Web Page : http://www.saurabhonline.org
           ----------------------------

anyMail is an anonymous mailer. You can use it to send email without logging into an account. It uses the Winsock control in VB. So you need Mswinsck.ocx to be in your system directory. I have included the source code for it. It uses SMTP relaying to send the message to any e-mail address from any email address. I am also including two lists of SMTP servers. Rename the list you want to use to "servers.txt". If you manually edit the server list, each entry should be in the format servername:port. While most of the servers in the list worked for me, you might find some servers require to login or do not relay messages anymore. I have also included RFC 0821 (Simple Mail Transfer Protocol) in case you are intrested in the SMTP protocol. Since this is the first release of anyMail I expect there to be a lot of bugs. Please feel free to write to me for bug reports or suggestions. I have made anyMail for informational purposes only, I will not responsible for the wrong use (if any) of it.

Ps: I have not included the compiled exe with this release. If you do not have visual studio on your computer write to me for the exe.