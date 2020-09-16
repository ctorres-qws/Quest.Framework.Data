<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production Last Week</title>
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
	currentDate = Date()
	LastDate = DateAdd("d", -7, currentDate)
	Monday = DateAdd("d", -((Weekday(currentDate) + 7 - 2) Mod 7), LastDate)
	Saturday = DateAdd("d", 6, Monday)
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM Y_INV WHERE DATEOut BETWEEN #" & Monday & "# AND #" & Saturday & "# Order BY WAREHOUSE, PART")
'DebugCode(Monday & "- " & Saturday & " - " & strSQL)
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
'rs2.Cursortype = 2
'rs2.Locktype = 3
'rs2.Open strSQL2, DBConnection
Set rs2 = GetDisconnectedRS(strSQL2, DBConnection)

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

        <ul id="Profiles" title="Production Last Week" selected="true">

<% 
rs.filter = "WAREHOUSE='WINDOW PRODUCTION'"

response.write "<li class='group'>WINDOW PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Prodcution Date</th><th>Floor / Notes </th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodlastweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Dateout") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='COM PRODUCTION' "

response.write "<li class='group'>COM PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Production Date</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodweektabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("DateOut") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
