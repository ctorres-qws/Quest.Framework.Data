<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->
<!--Feb 2019 - Added USA View -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production This Week</title>
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
Monday = DateAdd("d", -((Weekday(currentDate) + 7 - 2) Mod 7), currentDate)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("SELECT * FROM Y_INV WHERE DATEOut BETWEEN #" & Monday & "# AND #" & currentDate & "# Order BY WAREHOUSE, PART", isSQLServer)
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
        <h1 id="pageTitle">Production This Week</h1>
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

        <ul id="Profiles" title="Stock in Production This Week" selected="true">
         <li class="group"><a href="productionweek.asp?part=<%response.write part%>" target="_self" >Production Today (Table Form) - Switch to Row Form</a></li>

<%

if CountryLocation = "USA" then

rs.filter = "WAREHOUSE='JUPITER PRODUCTION'"
Response.write "<li class='group'>--------------" & MONDAY & " - " & currentDate & " --------------</li>"
Response.write "<li class='group'>JUPITER PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Production Date</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
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
Loop
Response.write "</table></li>"


else

rs.filter = "WAREHOUSE='WINDOW PRODUCTION'"
Response.write "<li class='group'>--------------" & MONDAY & " - " & currentDate & " --------------</li>"
Response.write "<li class='group'>WINDOW PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Production Date</th><th>Floor / Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
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
Loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='COM PRODUCTION' "

Response.write "<li class='group'>COM PRODUCTION</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Production Date</th><th>Notes</th><th>Allocation</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"

Do While Not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	If rs2.eof Then
		Description = "N/A"
	Else
		Description = rs2("Description")
	End If

	Response.write "<tr>"
	Response.write "<td><a href='stockbyrackedit.asp?ticket=prodweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
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
Loop

Response.write "</table></li>"


end if'canada/USA

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
