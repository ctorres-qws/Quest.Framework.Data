<!--#include file="dbpath_Quest_ArchiveLists.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			
			<!--CutlistArchiveProcedureMain.asp - is the Email code for sending this report-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest.mdb Archive Procedure</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>


<!--#include file="todayandyesterday.asp"-->
<% 
Server.ScriptTimeout=500
	currentDate = Date
	weekNumber = DatePart("ww", currentDate)
	OneWeekAgo = DateAdd("d",-7,currentDate)
	TwoWeekAgo = DateAdd("d",-14,currentDate)
	FourWeekAgo = DateAdd("d",-28,currentDate)
	CheckMinDate = DateAdd("yyyy",-5,currentDate)

 
'Collect TableNames from Schema Table 
Const adSchemaTables = 20
Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")

%>
</head>
<body>
<ul id="screen1" title="Quest Dashboard" selected="true">

	<li><b><u>CutList Archive Status: <%Response.write currentDate%> </u></b></li>
	<li><p>
	This email reflects archive update information from the last week as the first portion of information flow.<BR>
	Please respond to this email to Ariel Aziza and Michael Bernholtz with any cutlist issues this week.<BR>
	Specifically any cut-lists that could not be cut on the machines and had to be cut manually instead.<BR>
	</p></li>
	<li><b><i>Archive 1 - CUT/HCUT/DMSAW/ROD</i></b></li>
		<!--#include file="cutlistArchiveProcedure1.asp"-->
	<li><b><i>Archive 2 - DMSDR/STOP/R3/Panel</i></b></li>
		<!--#include file="cutlistArchiveProcedure2.asp"-->
	<li><b><i>Archive 3 - SCRN</i></b></li>
		<!--#include file="cutlistArchiveProcedure3.asp"-->
	<li><b><i>Delete 4 - QSU/QSP/PANELS</i></b></li>
		<!--#include file="cutlistArchiveProcedure4.asp"-->
		
</ul>
  
<% 

rs.close
set rs=nothing
DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
%>


</body>
</html>
