<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->
<!-- May 2019 - Updated to include Texas Database-->

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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIPPING_TRUCK WHERE [Active] = TRUE ORDER BY ID ASC")
rs.Cursortype = 2
rs.Locktype = 3

if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if
%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle">Open Trucks</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "indexTexas.html#_Report"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Report"
				HomeSiteSuffix = ""
			end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
   
 
        <ul id="Profiles" title="Glass Report - Commercial" selected="true">
        <li>Shipping Active Trucks</li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Truck Number</th><th>Dock Number</th><th>Truck ID</th><th>Truck Name</th><th>Item count</th><th>ShipDate</th><th>View</th><th>Compare</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TruckNum") & "</td>"
	response.write "<td>" & RS("DockNum") & "</td>"
	response.write "<td>" & RS("ID") & "</td>"
	response.write "<td>" & RS("truckName") & "</td>"
	
	Counter = 0
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_SHIPPING WHERE [Truck] = " & RS("ID") & " ORDER BY TAG ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	if CountryLocation = "USA" then
		rs2.Open strSQL2, DBConnection_Texas
	else	
		rs2.Open strSQL2, DBConnection
	end if
	do while not rs2.eof
			Counter = Counter + 1
	rs2.movenext
	loop
	rs("Itemcount") = Counter
	rs.update
	rs2.close
	set rs2 = nothing
	
	response.write "<td>" & Counter & "</td>"
	response.write "<td>" & RS("ShipDate") & "</td>"
	response.write "<td><a class='greenButton' href='ShippingTruckViewer.asp?truck=" & RS("ID") & "&Ticket=Open' target='_self' >View All Items </a></td>"
	response.write "<td><a class='greenButton' href='ShippingTruckCompare.asp?Job=" & RS("Job") & "&Floor=" & RS("Floor") & "&Ticket=Open' target='_self' >Check for Missing </a></td>"
	response.write "</tr>"
	rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
               
    </ul>      
  
</body>
</html>
