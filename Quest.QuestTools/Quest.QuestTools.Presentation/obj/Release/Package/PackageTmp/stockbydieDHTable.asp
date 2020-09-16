<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		 <!--Created January 2015 to be both Table and Row form to match StockLevels Durapaint and Horner-->		
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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'TILTON'  OR WAREHOUSE = 'TORBRAM' ORDER BY AISLE, RACK, SHELF ASC"
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
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
               <a class="button leftButton" type="cancel" href="stock2.asp" target="_self">Stock by Die</a>
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
    <li class="group"><a href="stockbydieDH.asp?part=<%response.write part%>" target="_self" >Stock (Table Form) - Switch to Row Form</a></li>

    <%
	
	
	rs.movefirst
rs.filter = "PART = '" & part & "'"

response.write "<li><img src='/partpic/" & part & ".png'/> - " & Description & "</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Thickness</th><th>Quantity (SL)</th><th>PO</th><th>Bundle</th><th>Length (mm)</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th><th>Allocated</th></tr>"
do while not rs.eof
po = rs("PO")
response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
	if isnull(rs.fields("Colour")) then 
	response.write rs.fields("Project")
	else
	response.write rs.fields("Colour")
	end if
	'Added Lmm (length in mm at Request of Ruslan - January 16, Michael Bernholtz
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
	
	
	response.write "<td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td><td>" & rs("Allocation") & "</td></tr>"

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

