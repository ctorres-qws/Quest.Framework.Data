<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Consumption report giving Todays Cycle cuts from the Horizontal Machines-->
<!-- Created December 10th, 2014 by Michael Bernholtz - Requested by Michael Angel-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Horizontal Cycle Cuts</title>
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
	




%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
        <ul id="Profiles" title="Horizontal Cuts: <% response.write DateAdd("d",-1,Date())%>" selected="true">
        <li>Horizontal Cut Report</li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
		<li>Ecowall - Ecowall Jamb Machine </li>
		
<% 
response.write "<li><table border='1' class='sortable'><tr><th>Cut Cycle</th><th>Cut Status</th><th>Start Date</th><th>End Date</th></tr>"
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM ProECOHor WHERE [StartDate] = #" & DateAdd("d",-1,Date()) & "# ORDER by JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JobNumber") & "</td>"
	response.write "<td>" & RS("CutStatus") &"</td>" 
	response.write "<td>" & RS("StartDate") &"</td>" 
	response.write "<td>" & RS("FinishDate") &"</td>" 
	response.write "</tr>"
	rs.movenext
loop
response.write "</table>"

rs.close
set rs = nothing
%>
<li>Q4750 - Q4750 Jamb Machine</li>
	
<% 
response.write "<li><table border='1' class='sortable'><tr><th>Cut Cycle</th><th>Cut Status</th><th>Start Date</th><th>End Date</th></tr>"
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM ProQHor WHERE [StartDate] = #" & DateAdd("d",-1,Date()) & "# ORDER by JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JobNumber") & "</td>"
	response.write "<td>" & RS("CutStatus") &"</td>" 
	response.write "<td>" & RS("StartDate") &"</td>" 
	response.write "<td>" & RS("FinishDate") &"</td>" 
	response.write "</tr>"
	rs.movenext
loop
response.write "</table>"

rs.close
set rs = nothing


' Close the Connection to Database
DBConnection.close 
set DBConnection = nothing
%>
               
    </ul>      
             
</body>
</html>
