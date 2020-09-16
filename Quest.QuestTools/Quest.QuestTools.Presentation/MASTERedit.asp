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

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title=" Master Part List - Edit" selected="true">
     <li class="group"><a href="mastereditTable.asp" target="_self" >Master Part List (Row Form) - Switch to Table Form</a></li>
    
    <li class="group">Extrusion List</li>
    <%
	
	rs.filter = "inventorytype = 'Extrusion'"

do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

	response.write "</a></li>"
rs.movenext
loop
%>
    <li class="group">Gasket List</li>
    <%
	
	rs.filter = "inventorytype = 'Gasket'"

do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

	response.write "</a></li>"

rs.movenext
loop

%>

    <li class="group">Hardware List</li>
    <%
	
	rs.filter = "inventorytype = 'Hardware'"

do while not rs.eof
	
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

	response.write "</a></li>"

rs.movenext
loop

%>

    <li class="group">Plastic List</li>
    <%
	
	rs.filter = "inventorytype = 'Plastic'"
	
do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

	response.write "</a></li>"

rs.movenext
loop

%>

    <li class="group">Sheet List</li>
    <%
	
	rs.filter = "inventorytype = 'Sheet'"


do while not rs.eof
	
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

	response.write "</a></li>"

rs.movenext
loop

%>

    <li class="group">All Other items List</li>
    <%
	
	rs.filter = "inventorytype = NULL"
	

do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
	response.write "<li><img src='/partpic/" & partfilename & ".png'/></li>"
	response.write "<li><a href='MASTEReditform.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " (" & rs("inventorytype") & ") " & rs("kgm") & " Kg/m"

response.write "</a></li>"

rs.movenext
loop




RESPONSE.WRITE "</UL>"


%>

</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

