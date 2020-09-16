<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--Updated January 2014 to be both Table and Row form-->	
<!-- February 2019 - Added USA option-->	 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
  
<% 
	
ticket = request.querystring("ticket")
	
Set rs = Server.CreateObject("adodb.recordset")
if CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'JUPITER' ORDER BY AISLE, RACK, SHELF ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'NPREP' ORDER BY AISLE, RACK, SHELF ASC"
end if
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

sortby = REQUEST.QueryString("sortby")


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
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
	<%	
	Select Case ticket
	Case "pic"
		%>
		<a class="button leftButton" type="cancel" href="stockbypic.asp" target="_self">Stock by Pic</a>
		<%
	Case Else
		%>
		<a class="button leftButton" type="cancel" href="stock2.asp" target="_self">Stock by Die</a>
		<%
	End Select
	%>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>
<%
	part = request.QueryString("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
		InventoryType = rs2("InventoryType")
	end if
%>	
<ul id="screen1" title="Stock by Die" selected="true">
    <!--Added Table form and Row Form option, Michael Bernholtz, January 2014-->
    <li class="group"><a href="stockbydie.asp?part=<%response.write part%>&ticket=<%response.write ticket%>" target="_self" >Stock (Table Form) - Switch to Row Form</a></li>

    <%
	
	if not rs.eof then
	rs.movefirst
	else
	end if 
rs.filter = "PART = '" & part & "' AND WAREHOUSE='GOREWAY'"

response.write "<li><img src='/partpic/" & part & ".png'/> - " & Description & " </li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Thickness</th><th>Quantity (SL)</th><th>PO</th><th>Bundle</th><th>Length (mm)</th><th>Aisle</th><th>Rack</th><th>Shelf</th></tr>"
do while not rs.eof
po = rs("PO")
response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
	if isnull(rs.fields("Colour")) then 
	response.write rs.fields("Project")
	else
	response.write rs.fields("Colour")
	end if
	response.write "</td>"
		Select Case InventoryType
	Case "Plastic"
		response.write "<td>" & rs.fields("Lft") & "</td>"
		response.write "<td></td>"
		response.write "<td>" & rs.fields("Qty") & "</td>"
		response.write "<td>" & po & "</td>"
		response.write "<td>" & rs.fields("Bundle") & "</td>"
		response.write "<td>" & rs.fields("Lmm") & "</td>"

	Case "Sheet"
		response.write "<td></td>"
		response.write "<td>" & rs.fields("Thickness") & "</td>"
		response.write "<td>" & rs.fields("Qty") & "</td>"
		response.write "<td>" & po & "</td>"
		response.write "<td>" & rs.fields("Bundle") & "</td>"
		response.write "<td></td>"
	Case Else 'Extrusion
		response.write "<td>" & rs.fields("Lft") & "</td>"
		response.write "<td></td>"
		response.write "<td>" & rs.fields("Qty") & "</td>"
		response.write "<td>" & po & "</td>"
		response.write "<td>" & rs.fields("Bundle") & "</td>"
		response.write "<td>" & rs.fields("Lmm") & "</td>"
		
	End Select
	
	
	response.write "<td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td></tr>"

rs.movenext
loop

RESPONSE.WRITE "</table></li>"
RESPONSE.WRITE "</ul>"

%>

</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

