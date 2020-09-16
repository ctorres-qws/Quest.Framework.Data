<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Spandrel Glass Colour-->
<!-- Created July 31st, by Michael Bernholtz -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Spandrel Colour Report</title>
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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_COLOR_SPANDREL ORDER BY CODE ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection



%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="SPANDREL GLASS - COLOUR" selected="true">
<% 
response.write "<li class='group'>Spandrel Colour REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Code</th><th>Description</th><th>Job</th><th>Notes</th><th>Active</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Code") & "</td><td>" & RS("Description") &"</td><td>" & RS("JOB") & "</td><td>" & RS("Notes") & "''</td><td>" & RS("Active") & "''</td></tr>"
	
	rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
