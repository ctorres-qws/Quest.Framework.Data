<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--#include file="dbpath.asp"-->
<!-- Created May 5th, by Michael Bernholtz at Request of Ariel Aziza -->
<!-- Stock by Colour list collected from stockcolorlist.asp-->

<!-- Switches back and forth to stockbycolortable-->

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
	Warehouse = Request.QueryString("Warehouse")
	If warehouse = "" or isNull(Warehouse) Then
		Warehouse = "GOREWAY"
	End If

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY AISLE, RACK, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Select Case warehouse
Case "ALL"
	Response.Write "<li> Inventory Items in All Warehouses of the Colour: " & colour & "</li>"
	rs.filter = "COLOUR = '" & colour & "'"
Case "ANP"
	Response.Write "<li> Inventory Items not in Production/Scrap of the Colour: " & colour & "</li>"
	rs.filter = "COLOUR = '" & colour & "' AND WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION' AND WAREHOUSE <> 'SCRAP'"
Case "GN"
	Response.Write "<li> Inventory Items in Goreway/Nashua of the Colour: " & colour & "</li>"
	rs.filter = "(WAREHOUSE = 'GOREWAY' AND COLOUR = '" & colour & "') OR (WAREHOUSE = 'NASHUA' AND COLOUR = '" & colour & "')"
Case Else
	Response.Write "<li> Inventory Items in: " & warehouse & " of the Colour: " & colour & "</li>"
	rs.filter = "COLOUR = '" & colour & "' AND WAREHOUSE = '" & warehouse & "'"
End Select

sortby = Request.QueryString("sortby")

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockcolorlist.asp" target="_self">Colour List</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Stock by Colour" selected="true">

<%
	colour = Request.QueryString("colour")
	Response.Write "<li class='group'><a href='stockbycolorTable.asp?colour=" & colour & "&warehouse=" & warehouse & "'' target='_self' >Stock (Row Form) - Switch to Table Form</a></li>"
	
rs.filter = "colour = '" & colour & "'"

Do While Not rs.eof
	po = rs("PO")
	Response.write "<li>" & rs.fields("part") & " "
	Response.write rs.fields("Colour")

	Response.write " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " PO" & po & " " & rs.fields("Lmm") & " mm" &"</li>"
	If warehouse = "GOREWAY" Then
		Response.write "<li>- Aisle " & rs.fields("Aisle") & " Rack " & rs.fields("rack") & " Shelf " & rs.fields("shelf") & "</li>"
	End If
'response.write "<li><a href='stockdet.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " " & rs.fields("Colour") & " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & "</a></li>"
'if color is missing put project
	rs.movenext
Loop

Response.WRITE "</UL>"

wcount=0
JFCHECKID=0
%>

</body>
</html>

<%
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>
