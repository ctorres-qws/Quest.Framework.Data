<!--#include file="dbpath.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- BARCODE SHIFT page changed from Direct Database call to V_REPORT3 call - Original code saved as BarcodeShiftBackup.asp-->
<!-- Michael Bernholtz, July 28th, at Request of Shaun Levy and Jody Cash -->
		 
		 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
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
strSQL = "Select * FROM V_Report3"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

totalg = 0
totalgp = 0
totala = 0
totalRedZ = 0
totalBlueZ = 0
totalForel = 0
totalWillian = 0
totalPanel = 0
totalAwning = 0
elist = ""

totalay = 0
totalgy = 0
totalgpy = 0
totalRedZy = 0
totalBlueZy = 0
totalForely = 0
totalWilliany = 0
totalPanely = 0
totalAwningy = 0

'totalc = 0
'totalsu = 0
'totalsp = 0

%>
<!--#include file="todayandyesterday.asp"-->
<%

rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear 

Do while not rs.eof
	totalg = totalg + rs("GLAZINGFULL")
	totalgp = totalgp + rs("GLAZINGPARTIAL")
	totala = totala + rs("ASSEMBLY")
	elist = elist + " " + rs("ERROR_EMPLOYEE")
	totalRedZ = totalRedZ + rs("ZipperRed")
	totalBlueZ = totalBlueZ + rs("ZipperBlue")
	totalForel = totalForel + rs("Glass_Forel")
	totalWillian = totalWillian + rs("Glass_Willian")
	totalPanel = totalPanel + rs("Panel")
	totalAwning = totalAwning + rs("Awning")


  rs.movenext
loop

rs.movefirst
rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYeary
Do while not rs.eof
	totalgy = totalgy + rs("GLAZINGFULL")
	totalgpy = totalgpy + rs("GLAZINGPARTIAL")
	totalay = totalay + rs("ASSEMBLY")
	elist = elist + " " + rs("ERROR_EMPLOYEE")
	totalRedZy = totalRedZy + rs("ZipperRed")
	totalBlueZy = totalBlueZy + rs("ZipperBlue")
	totalForely = totalForely + rs("Glass_Forel")
	totalWilliany = totalWilliany + rs("Glass_Willian")
	totalPanely = totalPanely + rs("Panel")
	totalAwningy = totalAwningy + rs("Awning")

  rs.movenext
loop

rs.close
set rs=nothing


DBConnection.close
set DBConnection=nothing
%>  


</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">
	<li class="group">This Week Stats</li>
		<li><% response.write "FULL GLAZING: " & totalg %></li>
		<li><% response.write "PARTIAL GLAZING: " & totalgp %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "Red Zipper: " & totalRedZ %></li>
		<li><% response.write "Blue Zipper: " & totalBlueZ %></li>
		<li><% response.write "Forel Glass: " & totalForel %></li>
		<li><% response.write "Willian Glass: " & totalWillian %></li>
		<li><% response.write "Panel: " & totalPanel %></li>
		<li><% response.write "Awning: " & totalAwning %></li>
		<li class="group">Last Week Stats</li>
		<li><% response.write "FULL GLAZING: " & totalgy %></li>
		<li><% response.write "PARTIAL GLAZING: " & totalgpy %></li>
		<li><% response.write "ASSEMBLY: " & totalay %></li>
		<li><% response.write "Red Zipper: " & totalRedZy %></li>
		<li><% response.write "Blue Zipper: " & totalBlueZy %></li>
		<li><% response.write "Forel Glass: " & totalForely %></li>
		<li><% response.write "Willian Glass: " & totalWilliany %></li>
		<li><% response.write "Panel: " & totalPanely %></li>
		<li><% response.write "Awning: " & totalAwningy %></li>
		<li class="group">Employee List</li>
		<li><% response.write "SYSTEM ERROR EMPLOYEES: " & elist %></li>
	</ul>

</body>
</html>

