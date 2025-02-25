<!--#include file="dbpath.asp"-->
<!-- Duplicate form of Stockedit.asp changed to organize by Colour instead of by Part-->   
<!-- Created january 17, 2014, requested by Ruslan, With permission by Jody Cash, by Michael Bernholtz-->  
<!-- Table Form of Stock by Colour edit Form September 2014, Michael Bernholtz-->   

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
'    SQL = "Select * FROM Y_INV ORDER BY Colour ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART, WAREHOUSE, COLOUR"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

sortby = REQUEST.QueryString("sortby")

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

colour = request.QueryString("colour")
%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stock.asp#_colour" target="_self">Stock by Colour</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Stock by Colour" selected="true">
    <!--Added Table form and Row Form option, Michael Bernholtz, September 2014-->
    <li class="group"><a href="stockeditcolour.asp?colour=<%response.write colour%>" target="_self" >Stock (Table Form) - Switch to Row Form</a></li>

    <li class="group">Stock</li>

<%

rs.movefirst
rs.filter = "COLOUR = '" & colour & "' AND WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION'"

Response.WRITE "<li> Inventory Items of the Colour: " & colour & "</li>"
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>Color</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Length (mm)</th><th>Warehouse</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>View Stock</th></tr>"

Do While Not rs.eof
	Response.write "<tr>"
	Response.write "<td>" & rs.fields("PART") & "</td>"
	Response.write "<td>" & rs.fields("COLOUR") & "</td>"
	Response.write "<td>" & rs.fields("LFT") & "</td>"
	Response.write "<td>" & rs.fields("QTY") & " SL</td>"
	Response.write "<td>" & rs.fields("PO") & "</td>"
	Response.write "<td>" & rs.fields("LMM") & "</td>"
	Response.write "<td>" & rs.fields("Warehouse") & "</td>"
	If rs.fields("warehouse") = "GOREWAY" then
		Response.write "<td>" & rs.fields("Aisle") & "</td>"
		Response.write "<td>" & rs.fields("Rack") & "</td>"
		Response.write "<td>" & rs.fields("Shelf") & "</td>"
	Else
		Response.write "<td></td><td></td><td></td>"
	End If

	Response.write "<td><a href='stockbyrackedit.asp?id=" & rs.fields("ID") & "&part=" & rs.fields("PART") & "&colour=" & rs.fields("colour") & "&ticket=colourtable' target='_self'>>View Stock </a></td>"
	Response.write "</tr>"
	rs.movenext
Loop

Response.WRITE "</table></li></UL>"

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing

%>

</body>
</html>
