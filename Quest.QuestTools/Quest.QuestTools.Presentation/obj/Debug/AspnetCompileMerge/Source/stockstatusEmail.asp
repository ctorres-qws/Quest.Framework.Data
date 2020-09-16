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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE [NOTE 2] = '*' ORDER BY WAREHOUSE, PART"
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
        <h1> Stock with Status Notes </h1>
        </div>

        <ul id="Profiles" title="Stock Entered Today" selected="true">

<% 
rs.filter = "WAREHOUSE='GOREWAY'"

Response.write "<br><b><u>GOREWAY</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='HORNER'"

Response.write "<br><b><u>HORNER</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='TILTON'"

Response.write "<br><b><u>TILTON</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"

Response.write "<br><b><u>HYDRO PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"



rs.filter = "WAREHOUSE='DURAPAINT'"

Response.write "<br><b><u>DURAPAINT PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"

Response.write "<br><b><u>DURAPAINT (WIP) PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='DEPENDABLE'"

Response.write "<br><b><u>DEPENDABLE PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='METRA'"

Response.write "<br><b><u>METRA PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"


rs.filter = "WAREHOUSE='EXTAL SEA'"

Response.write "<br><b><u>EXTAL PENDING</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='NASHUA'"

Response.write "<br><b><u>NASHUA</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"

rs.filter = "WAREHOUSE='NPREP'"

Response.write "<br><b><u>NASHUA PREP</u></b>"
Response.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	Response.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
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
	Response.write "<td>" & rs("Note 2") & " </td>"
	Response.write "</tr>"

	rs.movenext
loop
Response.write "</table>"


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
