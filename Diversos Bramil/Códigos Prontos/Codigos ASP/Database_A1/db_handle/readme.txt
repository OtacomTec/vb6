Installation
1 - copy the file in a Web Server directory (this directory must have execute script attribute)
2 - create an Access System DSN named "Instant" (select the file "instant.mdb"
as "database file"; if you have Access97 select the file "instant97.mdb";
if you want you can save this file in a different directory on your disk rather than
the Web directory where the other files were saved)

Custom
1 - open the file "main.cls" and change the value of MainBack to modify the back URL page from
"Main.asp".

Use
1 - access the page "Main.asp"

Handle Main Menu
1 - from Main.asp page select "1 - dB Command"
2 - select "1 - Command Structure"
3 - modify, add or delete record. In this way you will handle the structure of the
main "dB Administration" Menu.

NB: "do not delete" the record named "1 - dB Command" because in this case You will
be not able to handle the dB via the Web pages. However you can modify the name.

Add a dB Table to handle
1 - from Main.asp page select "1 - dB Command"
2 - select "2 - Sub Commands"
3 - select "Add" button. The "Add window" will open.
	= > MainID : identifies the Sub Command Menu
	= > Menu Title : Sub Command Menu voice
	= > DSN_Origin : insert here the DSN connection string. You can use an ODBC connection
		  	 (like the "Instant" voice defined to access this database) or an
			 embedded connection string (like: "PROVIDER= ...." or "DRIVER=...")
	= > Table : 	 the name of the table you want to admin. It can be a table or a query.
	= > SQL_Query :  (optional) you can insert a SQL query to view, for instance, only some
			 fields (NB: however to add/modify/delete records the system will use
			 the table name you have written in the "Table" field)
	= > Prm_Key : 	 the Primary Key Field name. If you leave blank this field you will
			 not be able to modify/delete records.
	= > Title :	 the Title of the window for this Table.
	= > Record_Cmd : you can insert here a value between 0 and 7:
				0 - no record commands
				1 - add new record
				2 - modify record
				3 - add/modify record
				4 - delete record
				5 - add/delete record
				6 - modify/delete record
				7 - add/modify/delete record
4 - Select confirm to add the record.

In this way tou will be able to handle the table you have added: simply choose the sub command
menu and the sub command menu voice.



			

	

