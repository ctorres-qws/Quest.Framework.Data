<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		<!--#include file="dbpath.asp"-->
<!-- Created Jan 6th, 2015 by Michael Bernholtz at Request of Shaun Levy-->
<!--Full list of Inventory Just for Viewing-->
<!-- February 2019 - Added USA option -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock By Warehouse</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
 <!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

  <script type="text/javascript">
	$(document).ready( function () {
		$('#Job').DataTable({
			"iDisplayLength": 25
		});
	});
  
  </script>
  <script type="text/javascript">
    iui.animOn = true;
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
	</head>
<body >

 <div class="toolbar">
        <h1 id="pageTitle">Warehouse View</h1>
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
<%
	warehouse = request.QueryString("warehouse")
	if warehouse <> "" then
		warehouse = replace(warehouse," + ", "&nbsp;")
	else 
		if CountryLocation = "USA" then
			warehouse = "JUPITER"
		else
			warehouse = "NASHUA"
		end if
	
	end if
%>	


<ul id="screen1" title="Stock by Warehouse" selected="true">

<li><form id="Warehouse" class="panel" name="Warehouse" action="stockbyWarehouseTable.asp" method="GET" target="_self" >

 <h2> Choose a location for inventory</h2>
<fieldset>

 <div class="row">
 
             <label>Warehouse</label>
            <select name="warehouse" onchange = "Warehouse.submit()">
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE ORDER BY NAME ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Select Case warehouse
Case ""
		rs2.movefirst
		warehouse = RS2("NAME")
Case "ALL"
	Response.Write "<option value='ALL' selected >ALL (CAN & USA)</option>"
Case "ANP"
	Response.Write "<option value='ANP' selected >ALL (No Production)(CAN & USA)</option>"
Case "NOW"
	Response.Write "<option value='NOW' selected >Current Inventory (No Pending)</option>"
Case "PEN"
	Response.Write "<option value='PEN' selected >Pending Items Only</option>"
Case ELSE
	rs2.filter = "NAME = '" & warehouse & "'"
	rs2.movefirst
	Response.Write "<option value='"
	Response.Write rs2("NAME")
	Response.Write "' selected >"
	Response.Write rs2("NAME")
	response.write ""
End Select



rs2.filter = ""
rs2.movefirst
Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""


rs2.movenext

loop
%>
<option value='ALL'>ALL</option>
<option value='ANP'>ALL (No Production)</option>
<option value='NOW'>Current Inventory (No Pending)</option>
<option value='PEN'>Pending Items Only</option>
</select></DIV>
</fieldset>
</form></li>




    <%
Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = 2
rs.Locktype = 3

	

Select Case warehouse
CASE "ALL"
	RESPONSE.WRITE "<li> Inventory Items in All Warehouse</li>"
	strSQL = "SELECT top 10000 * FROM Y_INV ORDER BY PART, ID DESC"
CASE "ANP"
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'JUPITER PRODUCTION' AND WAREHOUSE <> 'SCRAP' ORDER BY PART, WAREHOUSE ASC"
	RESPONSE.WRITE "<li> Inventory Items not in Production/Scrap</li>"
CASE "NOW"
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'JUPITER' OR  WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'NPREP' ORDER BY PART, WAREHOUSE ASC"
	RESPONSE.WRITE "<li> Current Inventory Items - Not including Pending </li>"
CASE "PEN"
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'SAPA' OR WAREHOUSE = 'HYDRO' OR WAREHOUSE = 'DEPENDABLE' OR  WAREHOUSE = 'DURAPAINT(WIP)' OR  WAREHOUSE = 'EXTAL SEA' OR  WAREHOUSE = 'CAN-ART' OR  WAREHOUSE = 'KEYMARK' OR  WAREHOUSE = 'METRA' ORDER BY PART, WAREHOUSE ASC"
	RESPONSE.WRITE "<li> Current Inventory Items - Not including Pending </li>"
CASE Else
	if warehouse = "WINDOW PRODUCTION" or warehouse = "JUPITER PRODUCTION" then
		OneMonth = DateAdd("m",-1,Date)
		strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = '" & warehouse & "' and DATEOUT > #" & OneMonth & "# ORDER BY PART, WAREHOUSE ASC"
		RESPONSE.WRITE "<li> Inventory Items in: " & warehouse & " over the last 30 days</li>"
	else
		strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = '" & warehouse & "'ORDER BY PART, WAREHOUSE ASC"
		RESPONSE.WRITE "<li> Inventory Items in: " & warehouse & "</li>"
	end if
	
End Select

rs.Open strSQL, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
'rs3.Cursortype = 2
'rs3.Locktype = 3
'rs3.Open strSQL3, DBConnection

Set rs3 = GetDisconnectedRS(strSQL3, DBConnection)


RESPONSE.WRITE "<li><table border='1' class='Job' id ='Job' style=' width: 100%'>"
RESPONSE.WRITE "<thead><tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Color PO</th><th>Bundle</th><th>Ex. Bundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th><th>Prod/Scrap</th><th>Prod Job</th><th>Prod Floor</th></tr></thead>"
RESPONSE.WRITE "<tbody>"

do while not rs.eof

if Left(rs.fields("PART"),1) = "." then
	part = "0" & rs.fields("PART")
else
	part = rs.fields("Part")
end if
	rs3.filter = "Part = '" & trim(part) & "'"
	if rs3.eof then 
		Description = "N/A"
	else
		Description = rs3("Description")
		InventoryType = rs3("InventoryType")
	end if

po = rs("PO")
response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
response.write rs.fields("Colour")

if InventoryType = "Sheet" then
response.write "</td><td> " & rs.fields("width") & " by " & rs.fields("height") & "</td>"
else
response.write "</td><td> " & rs.fields("Lft") & "'</td>"
end if

response.write "<td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td>" &  rs.fields("ColorPO") & " </td><td style='word-break:break-all;'> " & rs.fields("Bundle") & "</td><td style='word-break:break-all;'> " & rs.fields("ExBundle") & "</td>"
response.write "<td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td><td> " & year(rs.fields("dateout")) & "</td><td> " & rs.fields("jobcomplete") & "</td><td> " & rs.fields("Note") & "</td></tr>"

rs.movenext
loop

RESPONSE.WRITE "</Tbody></table></li>"

rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing

DBConnection.close
set DBConnection=nothing

%>


</ul>
</body>
</html>

