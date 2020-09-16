<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<!--Cleaned up July 15th - Opens many unecessary connections to the database - Michael Bernholtz -->

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

'Create a Query
    SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)

totalg = 0
totalgy = 0
totalgy2= 0
totalg2 = 0
totalg2y = 0
totalg2y2= 0
totala = 0
totalay = 0
totalay2= 0

%>
<!--#include file="todayandyesterday.asp"-->
<%
lastweek = weekNumber - 1
lastweek2 = weekNumber - 2


rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalg = totalg + 1
end if

IF rs("DEPT") = "GLAZING2" then
totalg2 = totalg2 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totala = totala + 1
end if

  rs.movenext
loop

rs.movefirst
rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy = totalgy + 1
end if

IF rs("DEPT") = "GLAZING2" then
totalg2y = totalg2y + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay = totalay + 1
end if

  rs.movenext
loop

'two week's prior stats
if rs.eof then
else
rs.movefirst
end if
rs.filter = "WEEK = " & lastweek2 & " AND YEAR = " & cYeary
Do while not rs.eof

IF rs("DEPT") = "GLAZING" then
totalgy2 = totalgy2 + 1
end if

IF rs("DEPT") = "GLAZING2" then
totalg2y2 = totalg2y2 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay2 = totalay2 + 1
end if

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


		<li class="group">This Week's Stats</li>
		<li><% response.write "GLAZING: " & totalg %></li>
		<li><% response.write "GLAZING2: " & totalg2 %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
        	<li class="group">Last Week's Stats</li>
		<li><% response.write "GLAZING: " & totalgy %></li>
		<li><% response.write "GLAZING2: " & totalg2y %></li>
		<li><% response.write "ASSEMBLY: " & totalay %></li>
             <li class="group">Two Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy2 %></li>
		<li><% response.write "GLAZING2: " & totalg2y2 %></li>
		<li><% response.write "ASSEMBLY: " & totalay2 %></li>
            
   
 

    </ul>
        
        
      





</body>
</html>



