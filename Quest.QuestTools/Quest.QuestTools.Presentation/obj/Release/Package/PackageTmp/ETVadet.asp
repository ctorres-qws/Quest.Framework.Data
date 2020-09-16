<!--#include file="dbpath.asp"-->
 <!--Updated May 2014 to prevent timeout-->                        
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
ticket = request.querystring("ticket")
employee = request.querystring("Employee")
first = request.querystring("first")
last = request.querystring("last")

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE WHERE EMPLOYEE = '" & employee & "' AND DEPT ='ASSEMBLY' AND YEAR = " & cYear
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	

Set rs1 = Server.CreateObject("adodb.recordset")
strSQL1 = "Select * FROM X_GLAZING WHERE EMPLOYEE = '" & employee & "' AND YEAR = " & cYear
rs1.Cursortype = 2
rs1.Locktype = 3
rs1.Open strSQL1, DBConnection
	
'Create a Query
    SQL3 = FixSQL("DELETE * FROM X_BARCODETEMPETV")
'Get a Record Set
    Set RS3 = DBConnection.Execute(SQL3)	
	
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODETEMPETV"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection



Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_EMPLOYEES"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
	<%	
	Select Case ticket
	Case "TodayAssembly"
		%>
		 <a class="button leftButton" type="cancel" href="ETVa.asp" target="_self">Assembly Today</a>

		<%
	Case "TodayGlazing"
		%>
		 <a class="button leftButton" type="cancel" href="ETVg.asp" target="_self">Glazing Today</a>
		<%
	Case "WeekAssembly"
		%>
		 <a class="button leftButton" type="cancel" href="ETVa.asp" target="_self">Assembly Week</a>

		<%
	Case "WeekGlazing"
		%>
		 <a class="button leftButton" type="cancel" href="ETVg.asp" target="_self">Glazing Week</a>
		<%
	Case "LastWeekAssembly"
		%>
		 <a class="button leftButton" type="cancel" href="ETVa.asp" target="_self">Assembly Last</a>

		<%
	Case "LastWeekGlazing"
		%>
		 <a class="button leftButton" type="cancel" href="ETVg.asp" target="_self">Glazing Last</a>
		<%
	Case else
		%>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
		<%
	End Select
		%>	
		
		
		
		
               
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Employee Details" selected="true">
    
    <li class="group">Detailed View</li>
    
    <%

employee = request.querystring("Employee")
first = request.querystring("first")
last = request.querystring("last")

response.write "<li><img src='/FACEBOOK/" & employee  & ".JPG' width='400' height='300'/ > " & last & ", " & first & "</li>"
' Facebook Folder is a Virtual Directory on IIS on 172.18.13.31. Folder must be updated if it changes (yearly)
'F:\projects\09 - PREF & Scanner Data\QUEST FACEBOOK\2015 EMPLOYEE PICS\CURRENT EMPLOYEES no names


rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth
rs1.filter = "DAY = " & cDay & " AND MONTH = " & cMonth

TOTALG = 0
TOTALG2 = 0
TOTALA = 0
Do while not rs.eof
	totalA = totalA + 1
rs.movenext
loop
Do while not rs1.eof
	if UCASE(rs1("FIRSTCOMPLETE")) = "TRUE" then
	TOTALG = TOTALG + 1
	else
	TOTALG2 = TOTALG2 + 1
	end if
rs1.movenext
loop

response.write "<li>Today: Assembly: " & totala & " - Glazing: " & totalg & " - Glazing2: " & totalg2 & "</li>"

rs.filter = "Week = " & weekNumber
rs1.filter = "Week = " & weekNumber
TOTALA = 0
Do while not rs.eof
	totalA = totalA + 1
rs.movenext
loop
Do while not rs1.eof
	if UCASE(rs1("FIRSTCOMPLETE")) = "TRUE" then
	TOTALG = TOTALG + 1
	else
	TOTALG2 = TOTALG2 + 1
	end if
rs1.movenext
loop


response.write "<li>Week: Assembly: " & totala & " - Glazing: " & totalg & " - Glazing2: " & totalg2 & "</li>"

rs.filter = "MONTH = " & cMonth
rs1.filter = "MONTH = " & cMonth
TOTALA = 0
Do while not rs.eof
	totalA = totalA + 1
rs.movenext
loop
Do while not rs1.eof
	if UCASE(rs1("FIRSTCOMPLETE")) = "TRUE" then
	TOTALG = TOTALG + 1
	else
	TOTALG2 = TOTALG2 + 1
	end if
rs1.movenext
loop


response.write "<li>Month: Assembly: " & totala & " - Glazing: " & totalg & " - Glazing2: " & totalg2 & "</li>"

response.write "<img src>"

RESPONSE.WRITE "</UL>"

%>
    
</body>
</html>



<% 

rs.close
set rs=nothing
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
rs5.close
set rs5=nothing
DBConnection.close
set DBConnection=nothing
%>

