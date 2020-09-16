<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Back Order Report - Showing all Back Order items as scanned -->
<!-- Michael Bernholtz, April 2015, Developed at Request of Jody Cash - Adapted mainly from Forel scan and XA9Backorder code-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Back Order Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target='_self' >Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Service Spandrel Glass" selected="true">
        
        

<li class='group'>Service Spandrel REPORT </li>
<li> Click on the Headers of each column to sort Ascending/Descending</li>  
<li><table border='1'class = 'Sortable' ><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Position</th><th>Type</th><th>PO</th><th>Po Line</th><th>Department</th><th> Date</th></tr>

    <%
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEGA WHERE DEPT = 'SP-SERVICE' ORDER BY ID DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & rs("JOB") & "</td>"
	response.write "<td>" & rs("FLOOR") & "</td>"
	response.write "<td>" & rs("Tag") & "</td>"
	response.write "<td>" & rs("POSITION") & "</td>"
	response.write "<td>" & rs("Type") & "</td>"
	response.write "<td>" & rs("PO") & "</td>"
	response.write "<td>" & rs("POLINE") & "</td>"
	response.write "<td>" & rs("DEPT") & "</td>"
	response.write "<td>" & rs("DATETIME") & "</td>"
	response.write "</tr>"
	rs.movenext
loop
response.write "</table></li></ul>"
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
