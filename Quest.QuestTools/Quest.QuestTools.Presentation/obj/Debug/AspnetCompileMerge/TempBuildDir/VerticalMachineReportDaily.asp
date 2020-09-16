                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Vertical Machine Report</title>
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
strSQL = "SELECT * FROM PROECOVERT WHERE STARTDATE = #" & Date() & "#  ORDER BY ID ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM PROECOVERT2 WHERE STARTDATE = #" & Date() & "#  ORDER BY ID ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Extrusion Report" selected="true">
        
        
<% 

VERT1 = 0
VERT2 = 0
response.write "<li class='group'>PROLINE ECOWALL VERTICAL 1</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Cycle</th><th>Cut Number</th><th>Cut Percentage</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JobNumber") & "</td>"
	response.write "<td>" & RS("CutNumber") & "</td>"
	response.write "<td>" & RS("CutStatus") &"%</td>"
	response.write " </tr>"
	VERT1 = VERT1 + RS("CUTNUMBER")
	rs.movenext
loop
Response.write "<tr><td><b>Total</b></td><td><b>" & VERT1 & "</b></td><td></td></th>"
response.write "</table></li>"

response.write "<li class='group'>PROLINE ECOWALL VERTICAL 2</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Cycle</th><th>Cut Number</th><th>Cut Percentage</th></tr>"
do while not rs2.eof
	response.write "<tr>"
	response.write "<td>" & RS2("JobNumber") & "</td>"
	response.write "<td>" & RS2("CutNumber") & "</td>"
	response.write "<td>" & RS2("CutStatus") &"%</td>"
	response.write " </tr>"
	VERT2 = VERT2 + RS2("CUTNUMBER")
	rs2.movenext
loop
Response.write "<tr><td><b>Total</b></td><td><b>" & VERT2 & "</b></td><td></td></th>"
response.write "</table></li>"

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing

%>
               
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
