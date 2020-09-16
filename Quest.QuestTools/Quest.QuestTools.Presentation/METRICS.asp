<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Detailed Stats</title>
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
    'SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    'Set RS = DBConnection.Execute(SQL)
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
	


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cYesterday = cDay - 1
cMonth = month(now)
cMonthy = cMonth
cYear = year(now)
cYeary = cYear
currentDate = Date
weekNumber = DatePart("ww", currentDate)
lastweek = weekNumber - 1

If cDay = 1 then
	if cMonth = 1 OR cMonth = 3 OR cMonth = 5 OR cMonth = 8 OR cMonth = 10 OR cMonth = 12 then
	cYesterday = 31
	end if
	if cMonth = 4 OR cMonth = 6 OR cMonth = 9 OR cMonth = 11 then
	cYesterday = 30
	end if
	if cMonth = 2 then
	cYesterday = 28
	end if
		
	if cMonth = 1 then
	cYeary = 12
	end if
	
	cMonthy = cMonth - 1

end if

totalo = 0
totalj = 0

'This metrics code fills out openings and joints for the Xbarcode table for one week
rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear ' & "TIME = 6:00:00 PM" 

Do while not rs.eof
tablename = rs("Job")
tag = rs("Tag")
floor = rs("Floor")

'response.write tablename
'response.write floor
'response.write tag
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From  [" & tablename & "] WHERE (((Floor) = '" & floor & "')) AND (((Tag) = '" & tag & "')) "
'strSQL = "SELECT * From  [" & tablename & "] "
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection


if rs2.bof then
rs("Last") = "A"
rs("First") = "A"
else
stylez = rs2("style")
'response.write rs("ID")

rs("Last") = Left(stylez,1)
rs("First") = (Left(stylez,1) + 1) * 2

	if rs("DEPT") = "GLAZING" then
	totalo = Left(stylez,1) + totalo
	end if
	
	if rs("DEPT") = "ASSEMBLY" then
	totalj = ((Left(stylez,1) + 1) * 2) + totalj
	end if

rs.update
end if
rs.movenext
loop

%>
</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

	<li class="group">Week's Stats</li>
    		<li><% response.write "OPENINGS: " & totalo %></li>
			<li><% response.write "JOINTS: " & totalj %></li>
	
        </ul>

</body>
</html>

<% 

rs.close
set rs=nothing


DBConnection.close
set DBConnection=nothing
%>

