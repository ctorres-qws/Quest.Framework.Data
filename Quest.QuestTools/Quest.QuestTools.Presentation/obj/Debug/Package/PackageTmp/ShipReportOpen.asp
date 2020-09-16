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
<!-- July 2019 New Format for sLIst Trucks-->

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
	   
<!-- DataTables CSS -->
	<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
	<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
	<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

	<script type="text/javascript">
		$(document).ready( function () {
		$('#Job').DataTable({
			"iDisplayLength": 25
		});
	});
  
	</script>	   


    <%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIP_TRUCK WHERE [Active] = TRUE ORDER BY DOCKNUM ASC")
rs.Cursortype = 2
rs.Locktype = 3

if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if
%>

    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle">Open Trucks</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "index.html#_Ship"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Ship"
				HomeSiteSuffix = ""
			end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
   
 
        <ul id="Profiles" title="Open Trucks" selected="true">
        <li>Shipping Active Trucks</li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
response.write "<li><table border='1' id='Job'><THead><tr><th>Truck Name</th><th>System Number</th><th>Jobs/Floors</th><th>Open Date</TH><th>Dock Number</th><th>Item count</th><th>View</th></tr></THead><TBody>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("truckName") & "</td>"
	response.write "<td>" & RS("ID") & "</td>"
	response.write "<td style='word-break:break-all;'>" & RS("sList") & "</td>"
	response.write "<td>" & RS("CreateDate") & "</td>"
	response.write "<td>" & RS("DockNum") & "</td>"
	
	Counter = 0
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND [Truck] = " & RS("ID") & " ORDER BY TAG ASC"
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
	response.write "<td><a class='greenButton' href='ShipTruckViewer.asp?truck=" & RS("ID") & "&Ticket=Open' target='_self' >View All Items </a></td>"
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
	</TBody></Table>
	</li>        
   </ul>      
  
</body>
</html>
