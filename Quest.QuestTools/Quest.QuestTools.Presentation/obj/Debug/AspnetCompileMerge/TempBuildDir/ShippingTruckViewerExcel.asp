<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Truck View</title>


    <%
	Truck = Request.QueryString("truck")
	Ticket = Request.QueryString("Ticket")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING WHERE [Truck] = " & Truck & " ORDER BY [TAG] ASC"

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=ScannedTruckData.xls"
  %>
 
        <ul id="Profiles" title="Window Report" selected="true">
<% 
do while not rs.eof
if rs("Window") = "Window" then 
Counter = Counter + 1
end if
rs.movenext
loop
rs.movefirst
response.write "<li> Total Windows " & Counter & "</li>"
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


%>
               
    </ul>      
  
</body>
</html>
