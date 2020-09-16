<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created January 2018, by Michael Bernholtz - View All items on All Trucks with the same name-->
<!-- Report page from ShippingTruckViewEnter.asp report -->
<!-- Sokol Requested the ability to see all items on multiple trucks that all have the same name-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Truck View</title>
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
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
		 <a class="button leftButton" type="cancel" href="ShippingTruckViewEnter.asp" target="_self">Shipping</a>
		
        </div>
        <ul id="Profiles" title="Window Report" selected="true">
		
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
Truck = Request.QueryString("Truck")
Counter = 0
NumTruck = 0
OldTruck = ""
NewTruck = ""
TruckList = ""
WindowNum = 0
Othernum = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING_Truck AS XST INNER JOIN X_SHIPPING AS XS ON XST.ID = XS.Truck WHERE XST.TruckName = '" & Truck & "' ORDER BY XST.ID DESC"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection

Set rs = GetDisconnectedRS(strSQL, DBConnection)

response.write "<li> Truck Name:" & Truck & "</li>"

response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Truck</th><th>Ship Date</th><th>Description</th></tr>"
rs.filter = ""
rs.movefirst
Do while not rs.eof
	OldTruck = NewTruck
	NewTruck = RS("Truck")
	if OldTruck <> NewTruck then
		NumTruck = NumTruck + 1
		TruckList =  RS("Truck")  & " - " & TruckList
	end if
	if RS("Window") = "Window" then
	WindowNum = WindowNum+1
	else
	OtherNum = OtherNum + 1
	end if

	response.write "<tr>"
		response.write "<td>" & RS("Job") & "</td>"
		response.write "<td>" & RS("Floor") & "</td>"
		response.write "<td>" & RS("TAG") & "</td>"
		response.write "<td>" & RS("Truck") & "</td>"
		response.write "<td>" & RS("Description") & "</td>"
		response.write "<td>" & RS("ShipDate") & "</td>"	
		response.write "</tr>"
	rs.movenext
Loop

Response.write "</table></li>"


response.write "<li> Number of Trucks " & NumTruck & "</li>"
response.write "<li> Truck Numbers " & TruckList & "</li>"
response.write "<li> Windows " & WindowNum & "</li>"
response.write "<li> Non Windows " & OtherNum & "</li>"


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
  
</body>
</html>
