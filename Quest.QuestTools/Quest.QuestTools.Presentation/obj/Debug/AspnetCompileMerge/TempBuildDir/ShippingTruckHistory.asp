<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- May 2017 - Report to show Last week worth of Closed Trucks -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping View</title>
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
Weekago = DateAdd("d",-7,Date)

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIPPING_TRUCK WHERE [Active] =  FALSE AND SHIPDATE > #" & Weekago & "# ORDER BY SHIPDATE DESC")
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
   
 
        <ul id="Profiles" title="Glass Report - Commercial" selected="true">
		<li>Click on the Headers of each column to sort Ascending/Descending</li>
        <li>Shipping Closed Trucks Since: <% response.write Weekago %></li>

	
<% 
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Truck Number</th><th>Dock Number</th><th>Truck Name</th><th>Item Count</th><th>OpenDate</th><th>ShipDate</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TruckNum") & "</td>"
	response.write "<td>" & RS("DockNum") & "</td>"
	response.write "<td>" & RS("truckName") & "</td>"
	response.write "<td>" & RS("itemCount") & "</td>"
	response.write "<td>" & RS("CreateDate") & "</td>"
	response.write "<td>" & RS("ShipDate") & "</td>"
	response.write "</tr>"
	rs.movenext
loop


rs.close
set rs = nothing


Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIPPING_TRUCK WHERE CREATEDATE > #" & Weekago & "# ORDER BY CreateDATE DESC")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
		</table></li>
		<BR>
		<BR>

        <li>Shipping Openned Trucks Since: <% response.write Weekago %></li>

	
<% 
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Truck Number</th><th>Dock Number</th><th>Truck Name</th><th>Item Count</th><th>OpenDate</th><th>ShipDate</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TruckNum") & "</td>"
	response.write "<td>" & RS("DockNum") & "</td>"
	response.write "<td>" & RS("truckName") & "</td>"
	response.write "<td>" & RS("itemCount") & "</td>"
	response.write "<td>" & RS("CreateDate") & "</td>"
	response.write "<td>" & RS("ShipDate") & "</td>"
	response.write "</tr>"
	rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
              </table></li> 
    </ul>      
       
</body>
</html>
