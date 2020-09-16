<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->
<!--May 2017 Jody requested change from Sort by Part to Sort by COLOUR and Add PRINT TO EXCEL button-->
<!--Feb 2019 - Added USA View -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production Today</title>
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
	CurrentDate = Request.Querystring("CDay")
	CDay = currentDate  
	If CDay = "" Then
		currentDate = Date()
		Yesterday = DateAdd("d",-1,Date())
	Else
		Yesterday = DateAdd("d",-1,CDay)
	End If

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("SELECT * FROM Y_INV WHERE DATEOUT = #" & currentDate & "# OR DATEOUT = #" & Yesterday & "# ORDER BY WAREHOUSE, Colour, PART",isSQLServer)
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Production Today</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>

        <ul id="Profiles" title="Production <% response.write CurrentDate %>" selected="true">
         <li class="group"><a href="productiontoday.asp" target="_self" >Production Today (Table Form) - Switch to Row Form</a></li>
	<% 

if CountryLocation = "USA" Then
else
%>	
		<li class="group"><a href="productiontodaytableexcel.asp" target="_self" >PRINT TO EXCEL</a></li>
<%
end if
%>
		
<% 

if CountryLocation = "USA" Then

rs.filter = "WAREHOUSE='JUPITER PRODUCTION' AND DATEOUT = #" & currentDate & "#"
Response.write "<li class='group'>--------------" & currentDate & " --------------</li>"
Response.write "<li class='group'>JUPITER PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=productiontodaytable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='JUPITER PRODUCTION' AND DATEOUT = #" & Yesterday & "#"
Response.write "<li class='group'>--------------" & Yesterday & " --------------</li>"
Response.write "<li class='group'> YESTERDAY JUPITER PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop

Response.write "</table></li>"
else

rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & currentDate & "#"

Response.write "<li class='group'>WINDOW PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=productiontodaytable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & currentDate & "#"

Response.write "<li class='group'>COM PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop

Response.write "</table></li>"

rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & Yesterday & "#"
Response.write "<li class='group'>--------------" & Yesterday & " --------------</li>"
Response.write "<li class='group'> YESTERDAY WINDOW PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop

Response.write "</table></li>"

rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & Yesterday & "#"

Response.write "<li class='group'>YESTERDAY COM PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes </th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
	Response.write "<td>" & Description & " </td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "<td>" & rs("Note") & " </td>"
	Response.write "<td>" & rs("Allocation") & " </td>"
	Response.write "<td>" & rs("Aisle") & " </td>"
	Response.write "<td>" & rs("Rack") & " </td>"
	Response.write "<td>" & rs("Shelf") & " </td>"
	Response.write "</tr>"

	rs.movenext
Loop

Response.write "</table></li>"

end if'Canada/USA

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
