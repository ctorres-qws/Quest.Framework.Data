<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stockpending.asp duplicated and put into table form, at Request of Ruslan Bedoev, May 23rd, 2014-->
<!-- Added USA View - February 2019-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Today</title>
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
strSQL = FixSQLCheck("SELECT * FROM Y_INV WHERE DATEIN BETWEEN #" & Monday & "# AND #" & currentDate & "# Order BY WAREHOUSE, PART",isSQLServer)
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
        <h1 id="pageTitle">Stock This Week</h1>
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

        <ul id="Profiles" title="Stock Entered This Week" selected="true">
         <li class="group"><a href="stockweek.asp?part=<%response.write part%>" target="_self" >Stock Today (Table Form) - Switch to Row Form</a></li>
<%

if CountryLocation = "USA" then

rs.filter = "WAREHOUSE='JUPITER'"

Response.write "<li class='group'>JUPITER</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

else



rs.filter = "WAREHOUSE='GOREWAY'"

Response.write "<li class='group'>GOREWAY</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='MILVAN'"

Response.write "<li class='group'>MILVAN</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"

Response.write "<li class='group'>HYDRO PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Allocation</th></tr>"

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='HORNER'"

Response.write "<li class='group'>HORNER</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='NASHUA'"

Response.write "<li class='group'>NASHUA</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='NPREP'"

Response.write "<li class='group'>NASHUA PREP</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DURAPAINT'"

Response.write "<li class='group'>DURAPAINT PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"

Response.write "<li class='group'>DURAPAINT (WIP) PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DEPENDABLE'"

Response.write "<li class='group'>DEPENDABLE PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Allocation</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='APEL'"

Response.write "<li class='group'>APEL PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Allocation</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='EXTAL SEA'"

Response.write "<li class='group'>EXTAL PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='EXTRUDEX'"

Response.write "<li class='group'>EXTRUDEX PENDING</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=inweektable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

end if' Canada/USA
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
