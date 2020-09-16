<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		<!--#include file="dbpath.asp"-->
<!-- Created May 5th, by Michael Bernholtz at Request of Ariel Aziza -->
<!-- Stock by Colour list collected from stockcolorlist.asp-->

<!-- Switches back and forth to stockbycolor.asp-->	 

<!-- New change - if not Goreway show external colour and allocation instead of Job Floor Tag - Michael Bernholtz requested by SHaun Levy Sept 2016-->
<!--Date: February 10, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add Milvan warehouse when displaying Aisle, Rack, Shelf
	
	Date: February 14, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add Milvan warehouse to Horner & Nashua drop-down and update report to include this stock
-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock By Colour</title>
 
<% 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=ColourTable" & Date() & ".xls"

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>

    </div>
<%
	colour = request.QueryString("colour")
	warehouse = request.QueryString("warehouse")
	currentDate = Now
	if warehouse <> "" then
		warehouse = replace(warehouse," + ", "&nbsp;")
	else
		warehouse = "ANP"
	end if
%>	


<ul id="screen1" title="Stock by Colour/Warehouse" selected="true">
    <!--Added Table form and Row Form option, Michael Bernholtz, January 2014-->
    


    <%
Set rs = Server.CreateObject("adodb.recordset")
	'strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	
	if CountryLocation = "USA" then
		WarehousesP = " AND (WAREHOUSE = 'JUPITER' OR WAREHOUSE = 'JUPITER PRODUCTION') "
		'WarehousesA = " AND (WAREHOUSE = 'JUPITER') "
	else
		WarehousesP = "AND (WAREHOUSE <> 'JUPITER' AND WAREHOUSE <> 'JUPITER PRODUCTION')"
		'WarehousesA = "AND (WAREHOUSE <> 'JUPITER')"
	end if



	Select Case warehouse
		Case "ALL"
			Response.Write "<li> Inventory Items in All Warehouses of the Colour: " & colour & "</li>"
			strSQL = "SELECT * FROM Y_INV WHERE COLOUR = '"& colour &"' " & WarehousesP & " {0} ORDER BY PART ASC"
			strSQLAlt = "SELECT * FROM Y_INV WHERE Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' {0} " & WarehousesP & "AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
		Case "ANP"
			Response.Write "<li> Inventory Items not in Production/Scrap of the Colour: " & colour & "</li>"
			strSQL = "SELECT * FROM Y_INV WHERE COLOUR = '"& colour &"' AND WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'SCRAP' " & WarehousesP & " {0} ORDER BY PART ASC"
			strSQLAlt = "SELECT * FROM Y_INV WHERE WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'SCRAP' {0} " & WarehousesP & " AND Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
		Case "GN"
			Response.Write "<li> Inventory Items in Goreway/Nashua of the Colour: " & colour & "</li>"
			strSQL = "SELECT * FROM Y_INV WHERE COLOUR = '"& colour &"' AND (WAREHOUSE = 'GOREWAY' OR WAREHOUSE ='NASHUA') {0} ORDER BY PART ASC"
			strSQLAlt = "SELECT * FROM Y_INV WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE ='NASHUA') AND Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' {0} AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
		Case "HN"
			Response.Write "<li> Inventory Items in Horner/Nashua/Milvan of the Colour: " & colour & "</li>"
			strSQL = "SELECT * FROM Y_INV WHERE COLOUR = '"& colour &"' AND (WAREHOUSE = 'HORNER' OR WAREHOUSE ='NASHUA' OR WAREHOUSE ='MILVAN') {0} ORDER BY PART ASC"
			strSQLAlt = "SELECT * FROM Y_INV WHERE (WAREHOUSE = 'HORNER' OR WAREHOUSE ='NASHUA' OR WAREHOUSE ='MILVAN') AND Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' {0} AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
		Case Else
			Response.Write "<li> Inventory Items in: " & warehouse & " of the Colour: " & colour & "</li>"
			strSQL = "SELECT * FROM Y_INV WHERE COLOUR = '"& colour &"' AND WAREHOUSE = '" & warehouse & "' {0} ORDER BY Aisle ASC, Rack ASC, Shelf ASC"  'PART ASC   '*** Changed Default Sort Order for Inventory Count Purposes
			strSQLAlt = "SELECT * FROM Y_INV WHERE WAREHOUSE = '" & warehouse & "' AND Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' {0} AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
	End Select

	If Request("PartNo") <> "" Then
		strSQL = Replace(strSQL, "{0}", " AND Part='" & Request("PartNo") & "' ")
		strSQLAlt = Replace(strSQLAlt, "{0}", " AND Part='" & Request("PartNo") & "' ")
	Else
		strSQL = Replace(strSQL, "{0}", "")
		strSQLAlt = Replace(strSQLAlt, "{0}", "")
	End If

rs.Open strSQL, DBConnection


RESPONSE.WRITE "<li><table border='1' class='sortable' style=' width: 100%'>"
if warehouse = "GOREWAY" or warehouse = "NASHUA" or warehouse = "GN" or warehouse = "HN" or warehouse = "MILVAN" then
	RESPONSE.WRITE "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Colour PO</th><th>Bundle</th><th>Ext Bundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th></tr>"
	do while not rs.eof
		po = rs("PO")
		response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
		response.write rs.fields("Colour")

			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if

		response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("Bundle") & "</td>"
		response.write "<td>" & rs.fields("ExBundle") & " </td><td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td></tr>"

	rs.movenext
	loop
else
	RESPONSE.WRITE "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th>"
	if WAREHOUSE = "WINDOW PRODUCTION" or Warehouse = "COM PRODUCTION" or Warehouse = "JUPITER PRODUCTION" then
		response.write "<th>Production Date</th><th>Production Job</th><th>Production Floor</th>"
	end if
	response.write "<th>Bundle</th><th>Ex Bundle</th><th>Allocation</th><th>Colour PO</th><th>Warehouse</th>"
	
	response.write "</tr>"
	do while not rs.eof
	po = rs("PO")
	response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
	response.write rs.fields("Colour")

			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if

	response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td>"
	
	if WAREHOUSE = "WINDOW PRODUCTION" or WAREHOUSE = "COM PRODUCTION"  or Warehouse = "JUPITER PRODUCTION" then
		response.write "<td>" & rs.fields("DateOut") & " </td><td>" & rs.fields("JobComplete") & " </td><td> " & rs.fields("Note") & " </td>"
	end if
	response.write "<td style='word-wrap: break-word'> " & rs.fields("Bundle") & "</td>"
	response.write "<td>" & rs.fields("ExBundle") & " </td><td> " & rs.fields("Allocation") & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("warehouse") & "</td>"


	response.write "</tr>"
	
rs.movenext
loop
end if


RESPONSE.WRITE "</table></li>"


%>
<li>
	<div>Alternate Inventory Sources:</div>
<%

	rs.Close
	rs.Open strSQLAlt, DBConnection

	Response.Write "<li><table border='1' class='sortable' style=' width: 100%'>"

	If warehouse = "GOREWAY" or warehouse = "NASHUA" or warehouse = "GN" or warehouse = "HN" or warehouse = "MILVAN" Then
		Response.Write "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Colour PO</th><th>Bundle</th><th>Ext Bundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th></tr>"
		Do While Not rs.eof
			po = rs("PO")
			Response.Write "<tr><td>" & rs.fields("PART") & "</td><td>"
			Response.Write rs.fields("Colour")
			
			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if

			Response.Write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("Bundle") & "</td>"
			Response.Write "<td>" & rs.fields("ExBundle") & " </td><td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td></tr>"

		rs.movenext
		Loop
	Else
		Response.Write "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th>"
		If WAREHOUSE = "WINDOW PRODUCTION" or Warehouse = "COM PRODUCTION" or Warehouse = "JUPITER PRODUCTION" Then
			Response.Write "<th>Production Date</th><th>Production Job</th><th>Production Floor</th>"
		End If
		Response.Write "<th>Bundle</th><th>Ex Bundle</th><th>Allocation</th><th>Colour PO</th><th>Warehouse</th>"

		Response.Write "</tr>"
		Do While not rs.eof
			po = rs("PO")
			Response.Write "<tr><td>" & rs.fields("PART") & "</td><td>"
			Response.Write rs.fields("Colour")
			
			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if

			Response.Write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td>"

			if WAREHOUSE = "WINDOW PRODUCTION" or WAREHOUSE = "COM PRODUCTION" or WAREHOUSE = "JUPITER PRODUCTION" then
				Response.Write "<td>" & rs.fields("DateOut") & " </td><td>" & rs.fields("JobComplete") & " </td><td> " & rs.fields("Note") & " </td>"
			end if
			Response.Write "<td style='word-wrap: break-word'> " & rs.fields("Bundle") & "</td>"
			Response.Write "<td>" & rs.fields("ExBundle") & " </td><td> " & rs.fields("Allocation") & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("warehouse") & "</td>"

			Response.Write "</tr>"

			rs.movenext
		Loop
	End If

	Response.Write "</table></li>"
%>

</ul>
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

