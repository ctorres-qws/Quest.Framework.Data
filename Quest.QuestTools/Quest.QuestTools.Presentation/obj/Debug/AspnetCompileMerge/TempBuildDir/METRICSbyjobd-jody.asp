<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Detailed Stats</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<!--<meta http-equiv="refresh" content="120" >-->
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


</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">



	<li class="group">Day's Stats</li>
    
    
<% 

'Create a Query
    'SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    'Set RS = DBConnection.Execute(SQL)
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE ORDER BY JOB, DEPT ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
<!--#include file="todayandyesterday.asp"-->
<%

totalo = 0
totalj = 0
loopo = 0
loopj = 0
loopg = 0
loopa = 0


'This metrics code fills out openings and joints for the Xbarcode table for one week
rs.filter = "DAY = " & cday & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear & " AND DEPT <> 'GLAZING2' "

lastjob = rs("Job")

Do while not rs.eof
tablename = rs("Job")
tag = rs("Tag")
floor = rs("Floor")
job = rs("job")
loopo = 0
loopj = 0
loopg = 0
loopa = 0
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
	rs("Last") = 0
	rs("First") = 0
		
	else
	stylez = rs2("style")
	'response.write rs("ID")
	
	rs("Last") = Left(stylez,1)
	rs("First") = (Left(stylez,1) + 1) * 2
	
		if rs("DEPT") = "GLAZING" then
		loopo = Left(stylez,1)
		loopg = 1
		end if
		
		if rs("DEPT") = "ASSEMBLY" then
		loopj = ((Left(stylez,1) + 1) * 2)
		loopa = 1
		end if
	
	'	if rs("DEPT") = "GLAZING" then
	'	totalo = rs("first") + totalo
	'	end if
	'	
	'	if rs("DEPT") = "ASSEMBLY" then
	'	totalj = rs("last") + totalj
	'	end if
	
	rs.update
end if

if job = lastjob then
	totalo = totalo + loopo
	totalj = totalj + loopj
		
	else
	
	response.write "<li>" & lastjob & " OPENINGS: " & totalo & " JOINTS: " & totalj & ", O/W: " & round((totalo/gtotal),1) & " J/W: " & round((totalj/atotal),1) &   "</li>"
	totalo = loopo +0
	totalj = loopj +0
end if

 lastjob = job
 
 gtotalo = gtotalo + loopo
gtotalj = gtotalj + loopj
gtotal = gtotal + loopg
atotal = atotal + loopa
 
rs.movenext


rs2.close
set rs2=nothing
loop

response.write "<li>" & lastjob & " OPENINGS: " & totalo & " JOINTS: " & totalj & ", O/W: " & round((totalo/gtotal),1) & " J/W: " & round((totalj/atotal),1) &    "</li>"
	totalo = 0
	totalj = 0
		%>
    <li class="group">Days's Total Stats</li>
    <%
	response.write "<li>TOTAL OPENINGS: " & gtotalo & " TOTAL JOINTS: " & gtotalj & " O/W: " & round((gtotalo/gtotal),1) & " J/W: " & round((gtotalj/atotal),1) &    "</li>"

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

