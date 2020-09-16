<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>PO Search Glass Report</title>
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
	
	QTFILE = Request.querystring("QTFILE")
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE [QTFILE] LIKE '%" & QTFILE & "%'  ORDER BY ID ASC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - PO Search" selected="true">
<% 
response.write "<li class='group'>PO Search GLASS REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Required Date</th><th>Ordered</th><th>Optima</th><th>Cut/Received Exterior</th><th>Cut/Received Interior</th><th>Sealed</th><th>Shipped</th><tr>"
do while not rs.eof


	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("REQUIREDDATE") & "</td>"
	response.write" <td>" & RS("INPUTDATE") & "</td>"
	response.write "<td>" & RS("OPTIMADATE") & "</td>"
	response.write "<td>" & RS("ExtReceived") & "</td>"
	response.write "<td>" & RS("IntReceived") & "</td>"
	response.write "<td>" & RS("COMPLETEDDATE") & "</td>"
	response.write "<td>" & RS("SHIPDATE") & "</td>"
	response.write " </tr>"

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
