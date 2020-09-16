<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->
		
		<!-- BARCODE SHIFTY page changed from Direct Database call to V_REPORT3 call - Original code saved as BarcodeShiftyBackup.asp-->
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

'added
totalay = 0
totalgpy = 0
totalgy = 0

totalg2 = 0
totalgp2 = 0
totala2 = 0

elist = ""

%>
<!--#include file="todayandyesterday.asp"-->
<%

if weekday(currentDate) = 2 then
	monday = 1
	rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYear & " AND DAY = " & Day((DateAdd("d",-2,currentDate)))

	if rs.eof then
	else
		Do while not rs.eof
			totalg2 = rs("GLAZINGFULL")
			
			totala2 = rs("ASSEMBLY")
			elist = elist + rs("ERROR_EMPLOYEE")
		
		rs.movenext
		loop
	end if

	rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYear & " AND DAY = " &  Day((DateAdd("d",-3,currentDate)))

else

	monday = 0
	rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear & " AND DAY = " & cYesterday
end if

if rs.eof then
else

	
	Do while not rs.eof

		totalgp2 = rs("GLAZINGPARTIAL")
		totalg = rs("GLAZINGFULLD")
		totala = rs("DAY_ASSEMBLY")
		totalgy = rs("GLAZINGFULLN")
		totalay = rs("NIGHT_ASSEMBLY")
		elist = elist + rs("ERROR_EMPLOYEE")
	rs.movenext
	loop
end if

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

	<% if monday = 1 then %>
	<li>MONDAY Stats show: Showing Last Friday & Saturday</li>
	<% end if %>
	<li class="group"><% if monday = 1 then response.write "Last Friday " end if%> Day Stats</li>
			<li><% response.write "GLAZING: " & totalg %></li>
			<li><% response.write "ASSEMBLY: " & totala %></li>
	<li class="group"><% if monday = 1 then response.write "Last Friday " end if%> Night Stats</li>
			<li><% response.write "GLAZING: " & totalgy %></li>
			<li><% response.write "ASSEMBLY: " & totalay %></li>

	<% if monday = 1 then %>
	<li class="group"> Last Saturday Stats</li>
			<li><% response.write "GLAZING: " & totalg2 %></li>
			<li><% response.write "ASSEMBLY: " & totala2 %></li>
			
	<% end if %>
	<li class="group">Employee List</li>
			<li><% response.write "Partial GLAZING: " & totalgp2 %></li>
			<li><% response.write "SYSTEM ERROR EMPLOYEES: " & elist %></li>

        </ul>
        
        
      



</body>
</html>



