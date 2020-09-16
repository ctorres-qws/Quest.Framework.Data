<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stockpending.asp duplicated and put into table form, at Request of Ruslan Bedoev, May 23rd, 2014-->
<!-- Updated December 2014 - Added Metra -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Pending Table</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	<script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    
    <%
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT yI.*,yM.Description FROM Y_INV yI LEFT JOIN y_Master yM ON yM.Part = yI.Part WHERE yI.WAREHOUSE <> 'WINDOW PRODUCTION' AND yI.WAREHOUSE <> 'SCRAP' ORDER BY yI.WAREHOUSE, yI.PART"
Set rs = GetDisconnectedRS(strSQL, DBConnection)

%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Pending Stock" selected="true">
         <li class="group"><a href="stockpending.asp?part=<%response.write part%>" target="_self" >Stock Pending (Table Form) - Switch to Row Form</a></li>
        
<% 
rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYRDO'"

response.write "<li class='group'>HYDRO PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"


do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"



rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"

response.write "<li class='group'>DURAPAINT(WIP) PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DEPENDABLE'"

response.write "<li class='group'>DEPENDABLE PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If
	
Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='EXTAL SEA'"

response.write "<li class='group'>EXTAL PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='KEYMARK'"

response.write "<li class='group'>KEYMARK PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='CAN-ART'"

response.write "<li class='group'>CAN-ART PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='APEL'"

response.write "<li class='group'>APEL PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"


do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='METRA'"

response.write "<li class='group'>METRA PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='EXTRUDEX'"

response.write "<li class='group'>EXTRUDEX PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Allocation</th><th>Order Date</th></tr>"
do while not rs.eof

	If rs("Description") & "" = "" Then
		Description = "N/A"
	Else
		Description = rs("Description")
	End If

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=ordertable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "<td>" & rs("Datein") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
