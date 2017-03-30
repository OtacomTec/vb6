________________________________________________________________ 

Crystal Decisions Technical Support - scr8_dist_expert.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp


PRODUCT VERSION

The Crystal Reports Distribution Expert is for use with Crystal Reports version 8 and higher (Developer Edition only).

________________________________________________________________ 

DESCRIPTION

Compiled Reports and Distribution Expert add-in for Crystal Reports 8:

This update enables Crystal Reports 8 or 8.5 to compile and distribute reports to users
who do not have the Crystal Reports 8 or 8.5 Designer. Compiling a report creates an executable of the report which can be previewed, printed, or exported outside of the Crystal Reports 8 or 8.5 Designer. 

The Distribution Expert creates a setup package of Crystal runtime files with reports which can be installed on computers without the Crystal Reports 8 or 8.5 Designer.  After running the update, the "Compile Report..." and "Report Distribution Expert..." menu items
are made available under the Report menu in the Crystal Reports 8 or 8.5 Designer.

Note:==========

This download has been made available for backwards compatiblity with previous
versions of Crystal Reports. This download includes no new features or fixes from previous
versions of Crystal Reports. Users are encouraged to distribute reports using the new Web Component Server which has many advantages over Compiled Reports. For information on using the Web Component Server, please refer to Chapter 17 of the User's Guide and the Web Reporting 
Administrator's Guide (web.pdf), both of which are located in the \Docs directory of the product CD.

Compiled Reports and the Report Distribution Expert are expected to be phased out in the
next release of Crystal Reports.

Compiled Reports and the Report Distribution Expert are ONLY supported when used with the Developer Edition of  Crystal Reports 8 or 8.5.  Althoug it may run properly if used with the Professional or Standard editions of Crystal  Reports, its use is a violation of the license agreement.

================

The following are known issues with this release of Compiled Reports and the Distribution 
Expert:

 - Clicking the 'Next>>' button under the '1.Options' tab in the Distribution Expert causes
   the Distribution Expert to analyze for files and then shutdown unexpectedly under 
   certain circumstances.

 - If Crystal Reports 8 is installed using a Network install, the Distribution Expert cannot
   be accessed through the 'Compile Report' box (click the 'Report' menu, then click 'Compile 
   Report...'). The Distribution Expert can only be accessed directly by clicking the 'Report' 
   menu, and then clicking 'Report Distribution Expert...'.

-  Using the Distribution Expert to distribute reports using ODBC Data Sources generates a 
   list of several ODBC files (i.e. Odbc32.dll) under the '3. Third Party Dlls' tab. These
   files should not be distributed with the Distribution Expert. ODBC core components should
   instead be installed using Microsoft Data Access Components (MDAC). The latest version 
   of MDAC can be found at www.microsoft.com/data.


________________________________________________________________ 

FILES

rdwiz.exe
readme.txt
Compiled_Reports.pdf 

________________________________________________________________ 

INSTALLATION


Note: ======

You must have the Crystal Reports 8 or 8.5 Designer installed before running rdwiz.exe

============

1. Extract all files from scr8_dist_expert.exe into a single folder, 

2. Run rdwiz.exe. Click Next to install the Report Distribution Expert.

3. Refer to the file Compiled_Reports.pdf For a Compiled Reports tutorial and Troubleshooting guide included in this download.  

This document contains information on the compile and distribution process as well as some limitations and an extensive Knowledge Base article listing to known issues.	
 

________________________________________________________________
Last updated on Feb 21, 2002
________________________________________________________________ 
