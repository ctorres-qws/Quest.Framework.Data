<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!--Email at Request of Jody Cash on BaseCamp if this report changes, change StockToday and StockTodayTable July 28th, 2014 -->
<!-- Email form is Daily Activity Gmail -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Today</title>

<%
	currentDate = Date()
	If isSQLServer Then
		currentDate = DateAdd("d",-29,Date)
	End If

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("SELECT * FROM Y_INV WHERE DATEIN = #" & currentDate & "# ORDER BY WAREHOUSE, PART", isSQLServer)

DebugMsg(strSQL)
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
        <h1> Stock Entered Into Database on <% Response.write Date() %> </h1>
        </div>

        <ul id="Profiles" title="Stock Entered Today" selected="true">

<% 
rs.filter = "WAREHOUSE='GOREWAY'"

response.write "<br><b><u>GOREWAY</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='HORNER'"

response.write "<br><b><u>HORNER</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"

	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write"</table>"

rs.filter = "WAREHOUSE='NASHUA'"

response.write "<br><b><u>NASHUA</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"

	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop

Response.write "</table>"

rs.filter = "WAREHOUSE='NPREP'"

response.write "<br><b><u>NASHUA PREP</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"

	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop

Response.write "</table>"

rs.filter = "WAREHOUSE='MILVAN'"

response.write "<br><b><u>MILVAN</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"

	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop

Response.write "</table>"

rs.filter = "WAREHOUSE='TILTON'"

response.write "<br><b><u>TILTON</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"

response.write "<br><b><u>HYDRO PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"



rs.filter = "WAREHOUSE='DURAPAINT'"

response.write "<br><b><u>DURAPAINT PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"


rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"

response.write "<br><b><u>DURAPAINT (WIP) PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='DEPENDABLE'"

response.write "<br><b><u>DEPENDABLE PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='APEL'"

response.write "<br><b><u>APEL PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"


rs.filter = "WAREHOUSE='METRA'"

response.write "<br><b><u>METRA PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"


rs.filter = "WAREHOUSE='EXTAL SEA'"

response.write "<br><b><u>EXTAL PENDING</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
Response.write "<tr>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & "</td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='EXTRUDEX'"

response.write "<br><b><u>EXTRUDEX</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"
end if

do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"

	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

	Response.write "<tr>"
	Response.write "<td>" & rs("part") & "</td>"
	Response.write "<td>" & Description & "</td>"
	Response.write "<td>" & rs("qty") & " </td>"
	Response.write "<td>" & rs("colour") & " </td>"
	Response.write "<td>" & rs("po") & " </td>"
	Response.write "<td>" & rs("Bundle") & " </td>"
	Response.write "<td>" & rs("Lft") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write"</table>"

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
