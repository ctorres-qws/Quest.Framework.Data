<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Reworked May 2016 to include X_GLAZING details Michael Bernholtz on behalf of Jody Cash And Shaun Levy-->
<!-- Copy and Paste from Weekly Code and just add cDay to the code-->
<!-- Assembly and Glazing Stats seperated out per Job -->
<!-- Glazing collects Joint number in X_GLAZING, so it does not need to access the Job Table like Assembly does -->

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
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>
<!--#include file="todayandyesterday.asp"-->
<ul id="screen1" title="Quest Dashboard" selected="true">
	<li class="group">Day's Stats</li>
    
    
<% 

'Create a Query
    'SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    'Set RS = DBConnection.Execute(SQL)
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE WHERE DAY = " & cDAY & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear & " AND DEPT ='ASSEMBLY' ORDER BY JOB ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs1 = Server.CreateObject("adodb.recordset")
strSQL1 = "Select * FROM X_GLAZING WHERE DAY = "  & cDAY & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear & " AND FirstComplete ='TRUE' ORDER BY JOB, Floor ASC"
rs1.Cursortype = 2
rs1.Locktype = 3
rs1.Open strSQL1, DBConnection

totalo = 0
totalj = 0
loopo = 0
loopj = 0
loopg = 0
loopa = 0
gtotalo = 0
gtotalj = 0



response.write "<li><table border ='1'>"
response.write "<tr><th>Job</th><th>Assembly</th><th>Openings</th><th>Openings per Window</th></tr>"
lastjob = rs("Job")

Do while not rs.eof
tablename = rs("Job")
tag = rs("Tag")
floor = rs("Floor")
job = rs("job")
loopo = 0
loopa = 0
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From  [" & tablename & "] WHERE (((Floor) = '" & floor & "')) AND (((Tag) = '" & tag & "')) "
'strSQL = "SELECT * From  [" & tablename & "] "
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL, DBConnection

if rs2.bof then
	rs("Last") = 0
	rs("First") = 0
		
	else
	stylez = rs2("style")
	'response.write rs("ID")
	
	rs("Last") = Left(stylez,1)
	rs("First") = (Left(stylez,1) + 1) * 2
	
		if rs("DEPT") = "ASSEMBLY" then
		loopo = Left(stylez,1)
		loopa = 1
		end if
		
	rs.update
end if

if job = lastjob then
	totalo = totalo + loopo
		
	else
	
	response.write "<tr><td>" & lastjob & "</td>" ' Job
	response.write "<td>" & aJob & "</td>" ' Assembly
	response.write "<td>" & totalo & "</td>"  ' Openings
	if aJob >0 then
		response.write "<td>" & round((totalo/aJob),1) & "</td>" ' Openings Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Openings Per Window
	end if
	
	

	aJob = 0
	totalo = loopo +0


end if

 lastjob = job
 
'These two variables add up each time an opening or joint is tallyed in the loop
gtotalo = gtotalo + loopo

'These totals look at total glazed or total assembled for aggregate use in a summary stat
atotal = atotal + loopa
aJob = aJob + loopa

 
 rs2.close
 set rs2 = nothing
rs.movenext
loop

response.write "<tr><td>" & lastjob & "</td>" ' Job
	response.write "<td>" & aJob & "</td>" ' Assembly
	response.write "<td>" & totalo & "</td>"  ' Openings
	if aJob >0 then
		response.write "<td>" & round((totalo/aJob),1) & "</td>" ' Openings Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Openings Per Window
	end if
	response.write "</table></Li>"

%>
<li class="group">Week's Total Stats</li>
    <%
	response.write "<li>TOTAL OPENINGS: " & gtotalo & "</li>"
	if atotal = 0 then
	response.write "<li>O/W: 0</li>" 
	else
	response.write "<li>O/W: " & round((gtotalo/atotal),1) & "</li>" 
	end if

	
response.write "<li><table border ='1'>"
response.write "<tr><th>Job</th><th>Glazing</th><th>Joints</th><th>Joints per Window</th></tr>"
lastjob = rs1("Job")

Do while not rs1.eof
tablename = rs1("Job")
tag = rs1("Tag")
floor = rs1("Floor")
job = rs1("job")
loopj = rs1("JOINTS")
loopg = 1
	

if job = lastjob then
	totalj = totalj + loopj
		
	else
	
	response.write "<tr><td>" & lastjob & "</td>" ' Job
	response.write "<td>" & gJob & "</td>" ' Glazing	
	response.write "<td>" & totalj & "</td>" ' Joints
	if gJob >0 then
		response.write "<td>" & round((totalj/gJob),1) & "</td>" ' Joints Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Joints Per Window
	end if


	gJob = 0
	totalj = loopj +0

end if

 lastjob = job
 
'These two variables add up each time an opening or joint is tallyed in the loop

gtotalj = gtotalj + loopj
'These totals look at total glazed or total assembled for aggregate use in a summary stat

gtotal = gtotal + loopg
gJob = gJob + loopg


rs1.movenext
loop



'response.write "<li>" & lastjob & " OPENINGS: " & totalo & " JOINTS: " & totalj & ", O/W: " & round((totalo/gtotal),1) & " J/W: " & round((totalj/atotal),1) &    "</li>"
	response.write "<tr><td>" & lastjob & "</td>" ' Job	
	response.write "<td>" & gJob & "</td>" ' Glazing	
	response.write "<td>" & totalj & "</td>" ' Joints
	if gJob >0 then
		response.write "<td>" & round((totalj/gJob),1) & "</td>" ' Joints Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Joints Per Window
	end if

	response.write "</table></li>"	

	
	
	totalo = 0
	totalj = 0
		%>
    <li class="group">Week's Total Stats</li>
    <%

	response.write "<li>TOTAL JOINTS: " & gtotalj & " </li>"
	if gtotal = 0 then
	response.write "<li>J/W: 0</li>" 
	else
	response.write "<li>J/W: " & round((gtotalj/gtotal),1) &    "</li>"
	end if



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

