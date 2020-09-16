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
	    


    <%
	Truck = Request.QueryString("truck")
	Ticket = Request.QueryString("Ticket")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING WHERE [Truck] = " & Truck & " ORDER BY TAG ASC"
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
        <h1 id="pageTitle"></h1>
		
		<%
		Select CASE Ticket
		Case "Open"
		%>
        <a class="button leftButton" type="cancel" href="ShippingReportOpen.asp" target="_self">Open</a>
		<%
		Case "Closed"
		%>
        <a class="button leftButton" type="cancel" href="ShippingReportClosed.asp" target="_self">Closed</a>
		<%
		Case Else
			if CountryLocation = "USA" then
		%>
			<a class="button leftButton" type="cancel" href="IndexTexas.html#_Ship" target="_self">Scan USA</a>
		<%
			Else
		%>
			<a class="button leftButton" type="cancel" href="Index.html#_Ship" target="_self">Scan Canada</a>
		<% 
			End if 
		End Select
		%>
        </div>
   
 
        <ul id="Profiles" title="Window Report" selected="true">
		<%
		if CountryLocation = "USA" then
		else
		%>
        <li><a href= "ShippingTruckViewerExcel.asp?Truck=<%response.write Truck %>&ticket=<%response.write Ticket %>" target="_self" >Send to Excel</a></li>
		<%
		end if%>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
do while not rs.eof
if rs("Window") = "Window" or rs("Window") = "H-Bar" then 
Counter = Counter + 1
end if
rs.movenext
loop
If Counter > 0 Then rs.movefirst
response.write "<li> Total Items   " & Counter & "</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th><th>Description</th></tr>"
do while not rs.eof

	response.write "<tr>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("SHIPDATE") & "</td>"
	response.write "<td>" & RS("Description") & "</td>"
	response.write "</tr>"

	rs.movenext
loop
response.write "</table></li>"


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
