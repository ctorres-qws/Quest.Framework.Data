<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
Job = Request.QueryString("Job")
Floor = Request.QueryString("Floor")
%>
  <title><%response.write JOB & " " &  Floor %></title>
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
		
	
        </div>
        <ul id="Profiles" title="Window Report" selected="true">
        <li><a href= "ShippingTruckFullReportExcel.asp?Job=<%response.write Job %>&FLoor=<%response.write Floor %>" target="_self" >Send to Excel (NOT READY)</a></li>
	
<% 
Job = Request.QueryString("Job")
	Floor = Request.QueryString("Floor")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING WHERE [Job] = '" & JOB & "' AND [Floor] = '" & Floor & "' ORDER BY TAG ASC"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection

Set rs = GetDisconnectedRS(strSQL, DBConnection)

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM " & JOB  & " WHERE [Floor] = '" & Floor & "' ORDER BY TAG ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection
MissingCounter = 0


do while not rs.eof
if rs("Window") = "Window" then 
Counter = Counter + 1
end if
rs.movenext
loop
If Counter > 0 Then rs.movefirst
response.write "<li>Table of Missing Windows</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Status</th></tr>"

Do while not RS2.eof
if Left(RS2("Tag"),1) = "-" then
JobTag = Right(RS2("TAg"),Len(RS2("TAg"))-1)
else
JobTag =RS2("Tag")
End if
rs.filter = "TAG = '" & JobTag & "'"
if rs.eof then
TruckTag = "Empty"
else
TruckTag = RS("Tag")
end if
if JobTag <> TruckTag then
MissingCounter = MissingCounter + 1

response.write "<tr>"
	response.write "<td>" & RS2("Job") & RS2("Floor") &  RS2("Tag") & "</td>"
	response.write "<td>" & RS2("Job") & "</td>"
	response.write "<td>" & RS2("Floor") & "</td>"
	response.write "<td>" & RS2("TAG") & "</td>"
	response.write "<td> Missing </td>"
	response.write "</tr>"

end if
rs2.movenext
loop


response.write "</table></li>"
response.write "<li> Total Windows Missing " & MissingCounter & "</li>"
response.write "<li> Total Windows Scanned for the Floor " & Counter & "</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Truck</th><th>Ship Date</th><th>Description</th></tr>"
rs.filter = ""

rs.movefirst

Do while not rs.eof
response.write "<tr>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("Truck") & "</td>"
	
	Set rsTruck = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT top 1 * FROM X_SHIPPING_TRUCK  WHERE [ID] = " & RS("TRUCK")
	rsTruck.Cursortype = 2
	rsTruck.Locktype = 3
	rsTruck.Open strSQL, DBConnection
	response.write "<td>" & rsTruck("SHipDate") & "</td>"
	rsTruck.Close
	Set rsTruck = nothing
	
	
	response.write "<td>" & RS("Description") & "</td>"
	response.write "</tr>"
rs.movenext
Loop


Response.write "</table></li>"


rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
  
</body>
</html>
