<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stocktoday.asp duplicated and put into table form, at Request of Ruslan Bedoev, May 23rd, 2014-->


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
	currentDate = Request.Querystring("CDay")
	CDay = currentDate  
	if currentDate = "" then
		currentDate = Date()
	End if
	
Set rs = Server.CreateObject("adodb.recordset")
If b_SQL_Server Then
	strSQL = "SELECT * FROM Y_INV WHERE ISNULL([NOTE 2],'') <> '' ORDER BY WAREHOUSE, PART"
Else
	strSQL = "SELECT * FROM Y_INV WHERE ISNULL([NOTE 2]) = FALSE and [NOTE 2] <> '' ORDER BY WAREHOUSE, PART"
End If
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
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Stock WIth Status Notes" selected="true">
         
<% 
rs.filter = "WAREHOUSE='GOREWAY'"
if not rs.eof then
response.write "<li class='group'>GOREWAY</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='HORNER'"
if not rs.eof then
response.write "<li class='group'>HORNER</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='TILTON'"
if not rs.eof then
response.write "<li class='group'>TILTON</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"
if not rs.eof then
response.write "<li class='group'>HYDRO PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th><th>Allocation</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DURAPAINT'"
if not rs.eof then
response.write "<li class='group'>DURAPAINT PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"
if not rs.eof then
response.write "<li class='group'>DURAPAINT (WIP) PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='DEPENDABLE'"
if not rs.eof then
response.write "<li class='group'>DEPENDABLE PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th><th>Allocation</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='METRA'"
if not rs.eof then
response.write "<li class='group'>METRA PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th><th>Allocation</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "<td>" & rs("Allocation") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='EXTAL SEA'"
if not rs.eof then
response.write "<li class='group'>EXTAL PENDING</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='NASHUA'"
if not rs.eof then
response.write "<li class='group'>NASHUA</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='NPREP'"
if not rs.eof then
response.write "<li class='group'>NASHUA PREP</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=statustable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"
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
