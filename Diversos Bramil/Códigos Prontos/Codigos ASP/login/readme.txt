Login Page by Sandman

This is a reasonable simple login system writen for asp aplications

it uses a database to store all the usernames and passwords
therefor it would be a good idea to password protect the database
i havent done it to the sample database to make life easier

there are 3 files included with this zip file (4 if you include this)

login.asp
reg.asp
db2.mdb

login.asp is the page that registered user can use to login to your 
page or aplication or whatever asp thing you want to password protect

reg.asp is used to regestier new users, you could add/remove fields to 
form to get more or less data from each regertering user, or you
could not use this page at all, meaning that only people that u put 
into the database would be able to use this

db2.mdb the database, its access2000 style so there you go

remember to change the path to database on both login.asp and reg.asp
to suit your server
the line below needs the part 'SOURCE=C:\Inetpub\wwwroot\test1\login\db2.mdb'
changed to mach your server

dataconn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\test1\login\db2.mdb"


USAGE

easiest way to use this to protect pages is to add this to the top of each
page you wish to protect

<% if session("logon") <> "yes" then response.redirect "logon.asp" %>

or something like that

if tried out pretty much everything i  can think of to test it with
and it all seems to work OK, but my ASP knolegde is limited (self taught)
so if you find something please let me know or ask in the forums