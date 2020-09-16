<!--#include file="dbpath.asp"-->
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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER ORDER BY Inventorytype, Part"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

sortby = REQUEST.QueryString("sortby")

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
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Master Part List - Edit" selected="true">
    <li class="group"><a href="masteredit.asp" target="_self" >Master Part List (Table Form) - Switch to Row Form</a></li>

    <li class="group"></li>
    
    <li class="group">Extrusion List</li>
<%

	rs.filter = "inventorytype = 'Extrusion'"

Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>KGM</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"
%>
    <li class="group">Gasket List</li>
<%

	rs.filter = "inventorytype = 'Gasket'"

Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>KGM</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "</tr>"
	
rs.movenext
loop
Response.write "</table></li>"

%>

    <li class="group">Hardware List</li>
<%

	rs.filter = "inventorytype = 'Hardware'"

Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>KGM</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "</tr>"
	
rs.movenext
loop
Response.write "</table></li>"

%>

    <li class="group">Plastic List</li>
    <%
	
	rs.filter = "inventorytype = 'Plastic'"
	
Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>LBF</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("LBF") & "</td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

%>

    <li class="group">Sheet List</li>
<%

	rs.filter = "inventorytype = 'Sheet'"

Response.WRITE "<li><table border='1' class='sortable'>"
Response.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>KGM</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

%>
    <li class="group">All Other items List</li>
<%

	rs.filter = "inventorytype = NULL"

RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>HYDRO</th><th>Description</th><th>Picture</th><th>KGM</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>Manage</th></tr>"

do while not rs.eof

Response.write "<tr>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>Manage</a></td>"
Response.write "<td> " & rs("CanArt") & "</td>"
Response.write "<td> " & rs("KeyMark") & "</td>"
Response.write "<td> " & rs("Extal") & "</td>"
Response.write "</tr>"

rs.movenext
loop
Response.write "</table></li>"

Response.WRITE "</UL>"

%>

</body>
</html>

<%

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>