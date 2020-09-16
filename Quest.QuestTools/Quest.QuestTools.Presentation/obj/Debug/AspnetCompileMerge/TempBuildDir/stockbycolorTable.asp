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
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  
  <script>
function ViewPartNo(str_PartNo) {
	document.Warehouse.PartNo.value = str_PartNo;
	document.Warehouse.submit();
}

function ReSetForm() {
	document.Warehouse.PartNo.value = "";
	document.Warehouse.Length.value = "";
	document.Warehouse.submit();
}

</script>


<% 

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
 <style>
  .csSearch { width: 90px !important; padding-left: 0px !important;}
 </style>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Stock By Color</h1>
			<a class="button leftButton" type="cancel" href="stockcolorlist.asp" target="_self">Color List</a>
    </div>
<%
	colour = request.QueryString("colour")
	warehouse = request.QueryString("warehouse")
	currentDate = Now
	If warehouse <> "" Then
		warehouse = replace(warehouse," + ", "&nbsp;")
	Else
		warehouse = "ANP"
	End If
%>

<ul id="screen1" title="Stock by Colour/Warehouse" selected="true">
<%
	if CountryLocation = "USA" then
	else
		%>
		<!-- Texas cannot do this 2019-->
		<li class='group'><a href='StockByColorTableExcel.asp?colour=<%Response.Write colour %>&warehouse=<%Response.Write warehouse %>&PartNo=<%= Request("PartNo") %>' target='_self'>Send to Excel</a></li>    
		<!--Added Table form and Row Form option, Michael Bernholtz, January 2014-->
		<li class="group"><a href="stockbycolor.asp?colour=<%Response.Write colour%>&warehouse=<%Response.Write warehouse%>" target="_self" >Stock (Table Form) - Switch to Row Form</a></li>
		<%
	end if
	%>
	<li><form id="Warehouse" class="panel" name="Warehouse" xaction="stockbycolorTable.asp" method="GET" target="_self" >
	<input type="hidden" name="colour" id= "colour" value = "<% Response.Write Colour%>">
	<h2> Choose a location for inventory</h2>
<fieldset>

<div class="row">

    <label>Warehouse</label>
    <select name="warehouse" onchange = "Warehouse.submit()">
<%

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM Y_WAREHOUSE " & Warehouses & " ORDER BY ID ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL2, DBConnection
			if CountryLocation = "USA" then
				rs2.filter = "Country = 'USA'"
			else
				rs2.filter = "Country = 'Canada'"
			end if
		Select Case warehouse
			Case ""
				rs2.movefirst	
				warehouse = RS2("NAME")
			Case "ALL"
				Response.Write "<option value='ALL' selected >ALL</option>"
			Case "ANP"
				Response.Write "<option value='ANP' selected >ALL (No Production)</option>"
			Case "GN"
				Response.Write "<option value='GN' selected >Goreway & Nashua</option>"
			Case "HN"
				Response.Write "<option value='HN' selected >Horner, Nashua & Milvan</option>"
			Case ELSE
				rs2.filter = "NAME = '" & warehouse & "'"
				rs2.movefirst
				Response.Write "<option value='"
				Response.Write rs2("NAME")
				Response.Write "' selected >"
				Response.Write rs2("NAME")
				Response.Write ""
		End Select
	
	if CountryLocation = "USA" then
		rs2.filter = ""
		rs2.filter = "Country = 'USA'"
	else
		rs2.filter = ""
		rs2.filter = "Country = 'Canada'"
	end if
		
	rs2.movefirst
	Do While Not rs2.eof

		Response.Write "<option value='"
		Response.Write rs2("NAME")
		Response.Write "'>"
		Response.Write rs2("NAME")
		Response.Write ""

		rs2.movenext

	Loop
%>
<option value='ALL'>ALL</option>
<option value='ANP'>ALL (No Production)</option>
<%
if CountryLocation = "USA" then
else
%>
<option value='GN'>Goreway & Nashua</option>
<option value='HN'>Horner, Nashua & Milvan</option>
<%
end if
%>

</select></DIV>
</fieldset>
</li>
<table>
	<tr>
		<td>&nbsp;&nbsp;</td><td>Part:&nbsp;</td><td><input type="text" name="PartNo" value="<%= Request("PartNo") %>"/></td><td>&nbsp;&nbsp;</td><td>Length:&nbsp;</td><td><input type="text" name="Length" value="<%= Request("Length") %>" class="csSearch"/></td><td>&nbsp;</td><td><input type="button" value="Search" onclick="Warehouse.submit()" style="padding: 5px 15px 5px 15px !important;"></td><td>&nbsp;</td><td><input type="button" value="Reset" onclick="ReSetForm()" style="padding: 5px 15px 5px 15px !important;"></td>
	</tr>
</table>

</form>

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
			strSQLAlt = "SELECT * FROM Y_INV WHERE WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'JUPITER PRODUCTION' AND WAREHOUSE <> 'SCRAP' {0} " & WarehousesP & " AND Colour in(SELECT Project FROM Y_Color WHERE Project <> '" & colour & "' AND Code IN(SELECT Code FROM y_color WHERE Project = '" & colour & "')) ORDER BY Colour ASC, PART ASC"
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

	Dim str_Where: str_Where = ""
	If Request("PartNo") <> "" Then
		str_Where = " AND Part='" & Request("PartNo") & "' "
		'strSQL = Replace(strSQL, "{0}", " AND Part='" & Request("PartNo") & "' ")
		'strSQLAlt = Replace(strSQLAlt, "{0}", " AND Part='" & Request("PartNo") & "' ")
	End If

	If Request("Length") <> "" Then
			str_Where = str_Where & " AND Lft=" & Request("Length") & " "
	End If

	If str_Where <> "" Then
		strSQL = Replace(strSQL, "{0}", str_Where)
		strSQLAlt = Replace(strSQLAlt, "{0}", str_Where)
	Else
		strSQL = Replace(strSQL, "{0}", "")
		strSQLAlt = Replace(strSQLAlt, "{0}", "")
	End If

	rs.Open strSQL, DBConnection

	Response.Write "<li><table border='1' class='sortable' style=' width: 100%'>"

	If warehouse = "GOREWAY" or warehouse = "NASHUA" or warehouse = "GN" or warehouse = "HN" or warehouse = "MILVAN" Then
		Response.Write "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Colour PO</th><th>Bundle</th><th>Ext Bundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th></tr>"
		Do While Not rs.eof
			po = rs("PO")
			'Response.Write "<tr><td>" & rs.fields("PART") & "</td><td>"
			Response.Write "<tr><td><a href='javascript: void();' onclick=""ViewPartNo('" & rs.fields("PART") &  "');"">" & rs.fields("PART") & "</a></td><td>"
			Response.Write rs.fields("Colour")
			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if
			Response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("Bundle") & "</td>"
			Response.Write "<td>" & rs.fields("ExBundle") & " </td><td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td></tr>"

		rs.movenext
		Loop
	Else
		Response.Write "<tr><th class='sorttable_alpha'>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th>"
		If WAREHOUSE = "WINDOW PRODUCTION" or Warehouse = "COM PRODUCTION" or Warehouse = "JUPITER PRODUCTION" Then
			Response.Write "<th>Production Date</th><th>Production Job</th><th>Production Floor</th>"
		End If
		Response.Write "<th>Bundle</th><th>Ex Bundle</th><th>Allocation</th><th>Colour PO</th><th>Warehouse</th>"

		Response.Write "</tr>"
		Do While not rs.eof
			po = rs("PO")
			'Response.Write "<tr><td>" & rs.fields("PART") & "</td><td>"
			Response.Write "<tr><td><a href='javascript: void();' onclick=""ViewPartNo('" & rs.fields("PART") &  "');"">" & rs.fields("PART") & "</a></td><td>"
			Response.Write rs.fields("Colour")

			if (RS.Fields("Width") + 0 > 0) then
				Response.Write "</td><td> " & rs.fields("Width") & " X " & rs.fields("Height") & "'</td>"
			else
				Response.Write "</td><td> " & rs.fields("Lft") & "'</td>"
			end if
			
			
			Response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td>"

			if WAREHOUSE = "WINDOW PRODUCTION" or WAREHOUSE = "COM PRODUCTION"  or Warehouse = "JUPITER PRODUCTION" then
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
			
			Response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td> " & rs.fields("colorpo") & "</td><td> " & rs.fields("Bundle") & "</td>"
			Response.Write "<td>" & rs.fields("ExBundle") & " </td><td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td></tr>"

		rs.movenext
		Loop
	Else
		Response.Write "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th>"
		If WAREHOUSE = "WINDOW PRODUCTION" or Warehouse = "COM PRODUCTION"  or Warehouse = "JUPITER PRODUCTION" Then
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
			
			Response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td>"

			if WAREHOUSE = "WINDOW PRODUCTION" or WAREHOUSE = "COM PRODUCTION" or Warehouse = "JUPITER PRODUCTION" then
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

