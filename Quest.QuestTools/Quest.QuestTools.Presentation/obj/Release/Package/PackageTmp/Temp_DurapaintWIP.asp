<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!-- All 2019 DURAPAINT WIP records-->
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Durapaint WIP 2019</title>
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
strSQL = "SELECT * FROM Y_INVLOG WHERE [YEAR] = 2019 AND WAREHOUSE = 'DURAPAINT(WIP)' AND Transaction <> 'original' ORDER BY PART ASC, ID ASC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>

        <ul id="Profiles" title="DURAPAINT(WIP)" selected="true">
<% 
response.write "<li class='group'>2019 WIP </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Day</th><th>Month</th><th>Year</th><th>Part</th><th>Qty</th><th>Len(Inch)</th><th>PO</th><th>BUNDLE</th><th>Color</th></tr>"
do while not rs.eof

		response.write "<TR>"
		Response.write "<td>" & RS("DAY") & "</td> "
		Response.write "<td>" & RS("MONTH") & "</td> "
		Response.write "<td>" & RS("YEAR") & "</td> "
		Response.write "<td>" & RS("PART") & "</td> "
		Response.write "<td>" & RS("QTY") & "</td> "
		Response.write "<td>" & RS("LINCH") & "</td> "
		Response.write "<td>" & RS("PO") & "</td> "
		Response.write "<td>" & RS("BUNDLE") & "</td> "
		Response.write "<td>" & RS("COLOUR") & "</td> "

		response.write " </tr>"
	
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
