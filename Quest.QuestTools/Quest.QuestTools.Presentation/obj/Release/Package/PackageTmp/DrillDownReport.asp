<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Glazing2 Reporting - Shows the Items cut based on JOB - Total, This Month, Today -->
<!-- Zipper Red, Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Drill Down Stats</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <!-- Request from Jody Cash on January 9th 2014 to change the Auto Refresh from 1200 to 90 -->
  <meta http-equiv="refresh" content="90" >
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

<!--#include file="todayandyesterday.asp"-->

</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
		<% 
			Ticket = Request.QueryString("Ticket") 
			If Ticket = "BarcoderTV" then
			BackButton = "BarcoderTV.asp"
			Else
			BackButton = "index.html#_Report"
			End if
		%>
                <a class="button leftButton" type="cancel" href="<%Response.Write BackButton %>" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Department Drill Down Stats" selected="true">

		<li class="group">Department Drill Down</li>
<%
DEPT = Request.QueryString("DEPT") 


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE WHERE DEPT = '" & DEPT & "' ORDER BY JOB ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear		
	
Response.write "<li class='group'>Today</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	Job1 =  UCASE(rs("Job"))
	Job2 = "0"
	JobCount = 0

	Do while not rs.eof
		Job2 = Job1
		Job1 = UCASE(rs("Job"))
		if Job1 = Job2 then
			JobCount = JobCount + 1
		else 
			response.write "<li>" & UCASE(Job2) & ": " & JobCount & "</li>"
			JobCount = 1
		end if
	rs.movenext
	loop
	response.write "<li>" & UCASE(Job1) & ": " & JobCount & "</li>"

end if

rs.filter = "Year = " & cYear & " AND Week = " & weekNumber 
	
Response.write "<li class='group'>This Week</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	Job1 =  UCASE(rs("Job"))
	Job2 = "0"
	JobCount = 0

	Do while not rs.eof
		Job2 = Job1
		Job1 = UCASE(rs("Job"))
		if Job1 = Job2 then
			JobCount = JobCount + 1
		else 
			response.write "<li>" & UCASE(Job2) & ": " & JobCount & "</li>"
			JobCount = 1
		end if
	rs.movenext
	loop
	response.write "<li>" & UCASE(Job1) & ": " & JobCount & "</li>"

end if

	
rs.filter = " Month = " & cMonth & " AND YEAR = " & cYear		

Job1 =  UCASE(rs("Job"))
Job2 = "0"
JobCount = 0
Response.write "<li class='group'>This Month</li>"	
Do while not rs.eof
	Job2 = Job1
	Job1 = UCASE(rs("Job"))
	if Job1 = Job2 then
		JobCount = JobCount + 1
	else 
		response.write "<li>" & UCASE(Job2) & ": " & JobCount & "</li>"
		JobCount = 1
	end if
rs.movenext
loop
response.write "<li>" & UCASE(Job1) & ": " & JobCount & "</li>"
%>

	</ul>
        
  
<% 
On Error Resume Next

rs.close
set rs=nothing
rs2.CLOSE
Set rs2= nothing
DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
