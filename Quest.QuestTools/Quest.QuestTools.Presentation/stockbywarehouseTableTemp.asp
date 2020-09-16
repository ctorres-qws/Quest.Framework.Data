<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		<!--#include file="dbpath.asp"-->
<!-- Created Jan 6th, 2015 by Michael Bernholtz at Request of Shaun Levy-->
<!--Full list of Inventory Just for Viewing-->


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
                <a class="button leftButton" type="cancel" href="index.html#_TmpINV" target="_self">TMP Stock</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>
<%
	warehouse = request.QueryString("warehouse")
	if warehouse <> "" then
		warehouse = replace(warehouse," + ", "&nbsp;")
	end if
%>	


<ul id="screen1" title="Stock by Warehouse" selected="true">

<li><form id="Warehouse" class="panel" name="Warehouse" action="stockbyWarehouseTableTEMP.asp" method="GET" target="_self" >

 <h2> Choose a location for inventory</h2>
<fieldset>

 <div class="row">
 
             <label>Warehouse</label>
            <select name="warehouse" onchange = "Warehouse.submit()">
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE ORDER BY ID ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Select Case warehouse
Case ""
		rs2.movefirst
		warehouse = RS2("NAME")
Case "ALL"
	Response.Write "<option value='ALL' selected >ALL</option>"
Case "ANP"
	Response.Write "<option value='ANP' selected >ALL (No Production)</option>"
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
strSQL = "SELECT * FROM Y_INVBU ORDER BY AISLE, RACK, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
	

Select Case warehouse
CASE "ALL"
	RESPONSE.WRITE "<li> Inventory Items in All Warehouse</li>"
CASE "ANP"
	rs.filter = " WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'SCRAP'"
	RESPONSE.WRITE "<li> Inventory Items not in Production/Scrap</li>"
CASE "NOW"
	rs.filter = " WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'HORNER' OR  WAREHOUSE = 'DURAPAINT'"
	RESPONSE.WRITE "<li> Current Inventory Items - Not including Pending </li>"
CASE "PEN"
	rs.filter = " WAREHOUSE = 'SAPA'  OR WAREHOUSE = 'HYDRO' OR WAREHOUSE = 'DEPENDABLE' OR  WAREHOUSE = 'DURAPAINT(WIP)' OR  WAREHOUSE = 'EXTAL SEA' OR  WAREHOUSE = 'CAN-ART' OR  WAREHOUSE = 'KEYMARK' OR  WAREHOUSE = 'METRA'"
	RESPONSE.WRITE "<li> Current Inventory Items - Not including Pending </li>"
CASE Else
	rs.filter = "WAREHOUSE = '" & warehouse & "'"
	RESPONSE.WRITE "<li> Inventory Items in: " & warehouse & "</li>"
End Select


RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Bundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Warehouse</th></tr>"
do while not rs.eof
po = rs("PO")
response.write "<tr><td>" & rs.fields("PART") & "</td><td>"
response.write rs.fields("Colour")

response.write "</td><td> " & rs.fields("Lft") & "'</td><td> " & " " & rs.fields("Qty") & " </td><td> PO" &  po & " </td><td style='word-break:break-all;'> " & rs.fields("Bundle") & "</td>"
response.write "<td>" & rs.fields("Aisle") & " </td><td> " & rs.fields("rack") & " </td><td> " & rs.fields("shelf") & "</td><td> " & rs.fields("warehouse") & "</td></tr>"

rs.movenext
loop

RESPONSE.WRITE "</table></li>"


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

