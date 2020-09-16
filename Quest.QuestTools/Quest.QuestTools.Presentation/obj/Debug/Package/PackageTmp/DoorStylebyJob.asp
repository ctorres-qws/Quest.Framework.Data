<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->
<!-- Date: June 14, 2019
	 Modified By: Michelle Dungo
	 Changes: Change keyword to use for displaying all values			  
-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>JPanel Colours</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
PARENT = Request.querystring("Parent")

gi_Mode = c_MODE_SQL_SERVER
Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select	

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
if UCASE(PARENT) = "ALLJOBS" then
	strSQL = "SELECT * FROM StylesDoor ORDER BY Job, NAME ASC"
else
	strSQL = "SELECT * FROM StylesDoor WHERE [Job] = '" & Parent & "' ORDER BY NAME ASC"
end if

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle"></h1>
			<a class="button leftButton" type="cancel" href="DoorStyle.asp" target="_self">New Search</a>
	</div>
    <ul id="Profiles" title="Door Styles - Job Search" selected="true">
<% 
response.write "<li class='group'>All Door Styles for Job - " & Parent & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Name</th><th>Job</th><th>Interior Door Type</th><th>Exterior Door Type</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Name") & "</td><td>" & RS("Job") & "</td><td>" & RS("IntDoorType") &"</td><td>" & RS("ExtDoorType") &"</td>"
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
End Function
%>
               
    </ul>                   
</body>
</html>
