<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Master Inventory Usage - Prepare by Michael Bernholtz for Vanessa Abraham and Mahesh Mohanlall February-->
<!-- For Every Master Item - Show most recent move to PRODUCTION -->
<!-- Original Code: Most Recent Receive Date and Most Recent Production Date -->
<!-- Updated Code: Most Recent Production Date only - Mahesh February 5th -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Master Inventory Usage</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

	<style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER ORDER BY ID ASC"
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
        <a class="button leftButton" type="cancel" href="index.html#_TmpINV" target="_self">Inv Count</a>
        </div>

		<form id="Complete" name="Complete"  method="GET" target="_self" selected="true" >  
        <ul id="Profiles" title="Glass Report - Commercial" selected="true">
        <li class='group'> Master Item Usage</li>
		<li>Report shows the most Recent Entry into Window Production by Each Master Iventory Item</li>

<% response.write "<li><table border='1' class='sortable'><tr><th>Master ID NAME</th><th>Description</th><th>KGM</th><th>LBF</th><th>Inventory Move Date</th><th>Inventory ID</th><th>PO</th><th>Bundle</th><th>Ex Bundle</th><th>JOB</th><th>Location</th></tr>"
do while not rs.eof
	
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT Top 1 * FROM Y_INV WHERE WAREHOUSE = 'WINDOW PRODUCTION' AND Part = '" & RS("PART") & "' ORDER BY DATEOUT DESC"
	rs2.Cursortype = GetDBCursorType
	rs2.Locktype = GetDBLockType
	rs2.Open strSQL2, DBConnection

		response.write "<tr>"
		response.write "<td>" & rs("Part") & "</td>"
		response.write "<td>" & rs("Description") & "</td>"
		response.write "<td>" & rs("KGM") & "</td>"
		response.write "<td>" & rs("LBF") & "</td>"
		
		if rs2.eof then
			response.write "<td>None</td><td></td><td></td><td></td><td></td><td></td><td></td>"
		else
			response.write "<td>" & rs2("Dateout") & " </td>"
			response.write "<td>" & rs2("ID") & "</td>"
			response.write "<td>" & rs2("PO") & "</td>"
			response.write "<td>" & rs2("Bundle") & "</td>"
			response.write "<td>" & rs2("ExBundle") & "</td>"
			response.write "<td>" & rs2("Allocation") & "</td>"
			response.write "<td>" & rs2("Warehouse") & "</td>"
		end if
		
		response.write " </tr>"
	
	rs2.close
	set rs2 = nothing	
	
	rs.movenext
loop
response.write "</table></ul>"

rs.close
set rs = nothing

DBConnection.close 
set DBConnection = nothing

%>

    </ul>
	<input type"hidden" id="ticket" name="ticket" value="Commercial" />
		</form>

</body>
</html>
