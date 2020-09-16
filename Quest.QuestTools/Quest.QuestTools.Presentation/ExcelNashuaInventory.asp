<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Import into Nashua Inventory into Excel for Inventory Counts -->
<!-- Table Values Organized by Aisle Rack Shelf-->
<!-- Nashua Inventory Requested by Shaun November 2018-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  
 <%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=NashuaInventory" & Date() & ".xls"
%>

<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'NASHUA' ORDER BY AISLE ASC, RACK ASC, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

    </head>
<body>
 <ul id="Profiles" title="Profiles" selected="true">

<% 

response.write "<li><table border='1' class='sortable'><tr><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Part</th><th>Quantity</th><th>PO</th><th>Bundle</th><th>External Bundle</th><th>Colour</th><th>Input Date</th><th>Size</th></tr>"
do while not rs.eof

	aisle = rs("aisle")
	rack = rs("rack")
	part = rs("part")
	qty = rs("qty")
	id = rs("ID")
	po = rs("PO")
	bundle = rs("Bundle")
	ExBundle = rs("ExBundle")
	width = rs("width")
	height = rs("height")
	shelf = rs("shelf")
	colour = rs("colour")
	datein = rs("datein")

%>
<tr>
<td><%response.write aisle %></td>
<td><%response.write rack %></td>
<td><%response.write shelf %></td>
<td><%response.write part %></td>
<td><%response.write qty %></td>
<td><%response.write PO %></td>
<td><%response.write bundle %></td>
<td><%response.write ExBundle %></td>
<td><%response.write colour %></td>
<td><%response.write datein %></td>

<td>
<%
if int(width) >1 then
response.write width & " by " & height 
else 
response.write " " 
end if 
%>

</td>

</tr>

<%

aisle = rs("aisle")
rack = rs("Rack")

rs.movenext
loop

rs.close
set rs= nothing
dbconnection.close
set dbconnection = nothing

%>
      </ul>
</body>
</html>
