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
  <title>Ship Floor View</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<meta http-equiv="refresh" content="1200" >
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	    


    <%
	Job = Request.QueryString("Job")
	Floor = Request.QueryString("Floor")
	Ticket = Request.QueryString("Ticket")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND [JOB] = '" & JOB & "' AND [Floor] = '" & Floor & "' ORDER BY TAG ASC"
rs.Cursortype = 2
rs.Locktype = 3
if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if

Set rs2 = Server.CreateObject("adodb.recordset")
if ticket = "search" then
strSQL2 = "SELECT * FROM X_SHIPPING_TRUCK WHERE [JOB] LIKE '%" & Job &  "%' AND [Floor] LIKE '%" & Floor &  "%'"
else
strSQL2 = "SELECT * FROM X_SHIP_TRUCK WHERE [sList] LIKE '%" & Job & Floor &  "%'"
end if

rs2.Cursortype = 2
rs2.Locktype = 3
if CountryLocation = "USA" then
	rs2.Open strSQL2, DBConnection_Texas
else	
	rs2.Open strSQL2, DBConnection
end if

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM [" & Job & "] WHERE [Floor] = '" & Floor &  "' ORDER BY TAG ASC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
<style>
table, th, td {
  padding: 10px;
}
th {
	font-size: 22px
}
</style>
    <div class="toolbar">
        <h1 id="pageTitle">Shipping Floor Report</h1>
		<%
		Select Case Ticket
		Case "open"
			%>
			<a class="button leftButton" type="cancel" href="ShipFloorOpen.asp" target="_self">Open Floors</a>
			<%
		Case "close"
			%>
			<a class="button leftButton" type="cancel" href="ShipFloorClosed.asp" target="_self">Closed Floors</a>
			<%
		Case "search"
			%>
			<a class="button leftButton" type="cancel" href="ShipFloorSearch.asp" target="_self">Search</a>
			<%
		Case else
			%>
			<a class="button leftButton" type="cancel" href="ShipFloorOpen.asp" target="_self">Open Floors</a>
			<%
		End Select

		%>
    </div>
   
        <ul id="Profiles" title="Window Report" selected="true">
		<li>Job: <% response.write Job %></li>
		<li>Floor: <% response.write Floor%></li>

		<% 
		Counter = 0
		rs2.filter = ""
		do while not rs2.eof
				Counter = Counter + 1
		rs2.movenext
		loop
		If Counter > 0 Then rs2.movefirst
		response.write "<li> # of Trucks Available:   " & Counter & "</li>"
		%>	
		
		<% 
		Counter = 0
		do while not rs.eof
			Counter = Counter + 1
		rs.movenext
		loop
		If Counter > 0 Then rs.movefirst
		response.write "<li> Total Items Scanned   " & Counter & "</li>"
		%>		
	
		<% 
		Counter = 0
		do while not rs3.eof
				Counter = Counter + 1
		rs3.movenext
		loop
		If Counter > 0 Then rs3.movefirst
		response.write "<li> Total Items Expected  " & Counter & "</li>"
		%>
			
			
			<table >
			<tr><TH>Trucks Available</TH><TH>Scanned Items</TH><TH>Back Order Items</TH></tr>
			<tr><td valign="Top">
<%
response.write "<table border='1' class='sortable' cellspacing='20' ><tr><th>Truck Name</th><th>Truck Number</th><th>Available Job/Floors</th><th>Create Date</th><th>Ship Date</th></tr>"
ListCount = 0
do while not rs2.eof
ListCount = ListCount + 1
	response.write "<tr>"
	response.write "<td>" & RS2("TruckName") & "</td>"
	response.write "<td>" & RS2("ID") & "</td>"
	response.write "<td>" & RS2("sList") & "</td>"
	response.write "<td>" & RS2("TAG") & "</td>"
	response.write "<td>" & RS2("CreateDate") &  "</td>"	
	response.write "<td>" & RS2("SHIPDATE") & "</td>"
	
	response.write "</tr>"

	rs2.movenext
loop

	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & ListCount & "</b></td>"
	response.write "</tr>"
	
response.write "</table>"
	



%>
</td><td valign="Top">
<%
response.write "<table border='1' class='sortable' cellspacing='20'><tr><th>Truck</th><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th><th>Description</th></tr>"

ListCount = 0
do while not rs.eof
ListCount = ListCount + 1
	response.write "<tr>"
	response.write "<td><b>" & RS("Truck") & "</b></td>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("SHIPDATE") & "</td>"
	response.write "<td>" & RS("Description") & "</td>"	
	response.write "</tr>"

	rs.movenext
loop

	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & ListCount & "</b></td>"
	response.write "</tr>"

response.write "</table>"
%>

</td><td valign="Top">
<%
response.write "<table border='1' class='sortable' cellspacing='20'><tr><th>Job</th><th>Floor</th><th>Tag</th></tr>"
ListCount = 0
TotalCount = 0
do while not rs3.eof
TotalCount = TotalCount + 1

	BarcodeTest = RS3("Job") & RS3("Floor") & RS3("TAG")
	rs.filter = " Barcode = '" & BarcodeTest & "'"
	if rs.eof then
		response.write "<tr>"
		response.write "<td>" & RS3("Job") & "</td>"
		response.write "<td>" & RS3("Floor") & "</td>"
		response.write "<td>" & RS3("TAG") & "</td>"	
	
		ListCount = ListCount + 1
	end if
	
	
	response.write "</tr>"

	rs3.movenext
loop

	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & ListCount & " / " & TotalCount & "</b></td>"
	response.write "</tr>"
	
	
response.write "</table>"
%>
</td><tr></table></li>


<%
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
               
    </ul>      
  
</body>
</html>
