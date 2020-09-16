<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Cycle Count</title>
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
strSQL = "SELECT * FROM Y_INV WHERE UCASE([Note 2]) = 'CC' AND (WAREHOUSE = 'JUPITER' OR WAREHOUSE = 'JUPITER PRODUCTION') ORDER BY AISLE ASC, RACK ASC, Shelf ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle"> Jupiter CC</h1>
		
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write HomeSite%>#_TmpINV" target="_self">INV-CC<%response.write HomeSiteSuffix%></a>

    </div>

        <ul id="Profiles" title="Inventory updated CC" selected="true">
<% 
response.write "<li class='group'>All JUPITER items Marked CC</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Part</th><th>Quantity</th><th>PO</th><th>Bundle</th><th>Ext. Bundle</th><th>Color</th><th>Enter Date</th><th>Modify Date</th><th>Warehouse</th><th>Note 2</th></tr>"
do while not rs.eof
		response.write "<tr><td>" & RS("Aisle") & "</td><td>" & RS("Rack") & "</td><td>" & RS("Shelf") &"</td><td>" & RS("Part") & "</td><td>" & RS("QTY") & "</td><td>" & RS("PO") & "</td><td>" & RS("Bundle") & "</td><td>" & RS("ExBundle") & "</td><td>" & RS("COLOUR") & "</td>" 
		response.write "<td>" & RS("DATEIN") & "</td><td>" & RS("ModifyDate") & "</td><td>" & RS("Warehouse") & "</td><td>" & RS("Note 2") & "</td>"
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
