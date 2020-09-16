<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!-- Page Created May 2015 - Information regulary requested by Shaun Levy after days off or Long Weekends -->
		<!--  V_REPORT3 call - Last 7 Days-->
		<!-- Michael Bernholtz, May 2015 -->

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
</script></head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>



<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM V_Report3"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


%>
<!--#include file="todayandyesterday.asp"-->

<ul id="screen1" title="Last 7 Days" selected="true">

	              <li class="group">LAST 7 Days</li>
				  
				 <li><table border='1' class='Job' id ='Job' ><thead><tr><th>Week Day</th><th>Date</th><th>Assembly</th><th>A Day</th><th>A Night</th><th>Full Glazing</th><th>G Day</th><th>G Night</th><th>Partial Glazing</th><th>Glazing SQFT</th><th>Errors</th><th>Red Zipper</th><th>Blue Zipper</th><th>Forel</th><th>Willian</th><th>All Panel</th><th>Awning Glazed</th></tr></thead><tbody>

<%

DaysAgo = 7
Do Until DaysAgo = 0

LoopDay = DAY(Date()- DaysAgo)
LoopWeek = DatePart("ww", Date()- DaysAgo)
LoopYear = YEAR(DATE() - DaysAgo)
LoopName = Weekday(DATE() - DaysAgo)
LoopWeekName = WeekDayName(LoopName) 
rs.filter = "WEEK = " & loopweek & " AND YEAR = " & loopyear & " AND DAY = " & LoopDay

response.write "<tr>"
response.write "<td>" & LoopWeekName & "</td>"
response.write "<td>" & Date()- DaysAgo & "</td>"
response.write "<td><b>" & rs("ASSEMBLY") & "</b></td>"
response.write "<td>" & rs("DAY_ASSEMBLY") & "</td>"
response.write "<td>" & rs("NIGHT_ASSEMBLY") & "</td>"
response.write "<td><b>" & rs("GlazingFull") & "</b></td>"
response.write "<td>" & rs("GlazingFullD") & "</td>"
response.write "<td>" & rs("GlazingFullN") & "</td>"
response.write "<td>" & rs("GlazingPartial") & "</td>"
response.write "<td><b>" & rs("SquareFoot") & "</b></td>"
response.write "<td>" & rs("ERROR_EMPLOYEE") & "</td>"
response.write "<td>" & rs("ZIPPERRED") & "</td>"
response.write "<td>" & rs("ZIPPERBLUE") & "</td>"
response.write "<td>" & rs("GLASS_FOREL") & "</td>"
response.write "<td>" & rs("GLASS_WILLIAN") & "</td>"
response.write "<td>" & rs("Panel") & "</td>"
response.write "<td>" & rs("Awning") & "</td>"

response.write "</tr>"

DaysAgo = DaysAgo - 1
loop

rs.close
set rs=nothing


DBConnection.close
set DBConnection=nothing
%>
		</table>
		</li>
        </ul>
        

</body>
</html>



