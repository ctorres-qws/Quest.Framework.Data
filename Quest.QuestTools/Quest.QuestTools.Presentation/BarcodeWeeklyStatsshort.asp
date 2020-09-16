<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


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
	
	
'Create a Query
    SQL3 = "DELETE * FROM X_BARCODETEMP1"
'Get a Record Set
    Set RS3 = DBConnection.Execute(SQL3)	
	
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODETEMP1"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_EMPLOYEES"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

'Set rs6 = Server.CreateObject("adodb.recordset")
'strSQL6 = "SELECT * From X_BARCODEGA ORDER BY DATETIME DESC"
'rs6.Cursortype = 2
'rs6.Locktype = 3
'rs6.Open strSQL6, DBConnection
	

JOB = REQUEST.QueryString("JOB")
FL = REQUEST.QueryString("FLOOR")

totalg = 0
totala = 0
totalc = 0
totalsu = 0
totalsp = 0


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
lastweek2 = weekNumber - 2
lastweek3 = weekNumber - 3
lastweek4 = weekNumber - 4
lastweek5 = weekNumber - 5
lastweek6 = weekNumber - 6
lastweek7 = weekNumber - 7
lastweek8 = weekNumber - 8
lastweek9 = weekNumber - 9
lastweek10 = weekNumber - 10


' Replacing old code (with errors) stored as backup in TodayandYesterday.inc, MIchael Bernholtz, January 2014
' If broken down into 4 parts - each with the months add by one for setting last month end
' April has 30 days, so May 1st sets - April 30, so if current month is may, set length of days in April
' Months with 31 (January, March, May, July, August, October, December)
' Months with 30 ( April, June, September, 3)
' February Leap years for 2016 until 2032 coded
' January starts a new Year
If cDay = 1 then
	if cMonth = 2 OR cMonth = 4 OR cMonth = 6 OR cMonth = 8 OR cMonth = 9 OR cMonth = 11 OR cMonth = 1 then
	cYesterday = 31
	end if
	if cMonth = 5 OR cMonth = 7 OR cMonth = 10 OR cMonth = 12 then
	cYesterday = 30
	end if
	if cMonth = 3 then
		if cyear = 2016 OR cyear = 2020 OR cyear = 2024 OR cyear = 2028 OR cyear = 2032  then 
			cYesterday = 29
		else
			cYesterday = 28
		end if
	end if
		
	cMonthy = cMonth - 1	
		
	if cMonth = 1 then
	cMonthy = 12
	cYeary = cYear - 1
	end if
	
	

end if

rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalg = totalg + 1
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

IF rs("DEPT") = "ASSEMBLY" then
totalay = totalay + 1
end if

  rs.movenext
loop

'two week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek2 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy2 = totalgy2 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay2 = totalay2 + 1
end if

  rs.movenext
loop
rs.filter =""

'three week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek3 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy3 = totalgy3 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay3 = totalay3 + 1
end if

  rs.movenext
loop
rs.filter =""

'four week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek4 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy4 = totalgy4 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay4 = totalay4 + 1
end if

  rs.movenext
loop
rs.filter =""

'five week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek5 & " AND YEAR = " & cYeary
Do while not rs.eof


IF rs("DEPT") = "GLAZING" then
totalgy5 = totalgy5 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay5 = totalay5 + 1
end if

  rs.movenext
loop
rs.filter =""

'six week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek6 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy6 = totalgy6 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay6 = totalay6 + 1
end if

  rs.movenext
loop
rs.filter =""

'seven week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek7 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy7 = totalgy7 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay7 = totalay7 + 1
end if

  rs.movenext
loop
rs.filter =""

'eight week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek8 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy8 = totalgy8 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay8 = totalay8 + 1
end if

  rs.movenext
loop
rs.filter =""

'nine week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek9 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy9 = totalgy9 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay9 = totalay9 + 1
end if

  rs.movenext
loop
rs.filter =""


'ten week's prior stats
rs.movefirst
rs.filter = "WEEK = " & lastweek10 & " AND YEAR = " & cYeary
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy10 = totalgy10 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay10 = totalay10 + 1
end if

  rs.movenext
loop
rs.filter =""

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
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "week: " & weeknumber %></li>
        	<li class="group">Last Week's Stats</li>
		<li><% response.write "GLAZING: " & totalgy %></li>
		<li><% response.write "ASSEMBLY: " & totalay %></li>
		<li><% response.write "week: " & lastweek %></li>
             <li class="group">Two Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy2 %></li>
		<li><% response.write "ASSEMBLY: " & totalay2 %></li>
		             <li class="group">Three Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy3 %></li>
		<li><% response.write "ASSEMBLY: " & totalay3 %></li>
		             <li class="group">Four Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy4 %></li>
		<li><% response.write "ASSEMBLY: " & totalay4 %></li>
		             <li class="group">FIve Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy5 %></li>
		<li><% response.write "ASSEMBLY: " & totalay5 %></li>
		<li><% response.write "week: " & lastweek5 %></li>
		             <li class="group">Six Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy6 %></li>
		<li><% response.write "ASSEMBLY: " & totalay6 %></li>
		             <li class="group">Seven Week's Prior Stats</li>
<!--		<li><% response.write "GLAZING: " & totalgy7 %></li>
		<li><% response.write "ASSEMBLY: " & totalay7 %></li>
		             <li class="group">Eight Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy8 %></li>
		<li><% response.write "ASSEMBLY: " & totalay8 %></li>
		             <li class="group">Nine Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy9 %></li>
		<li><% response.write "ASSEMBLY: " & totalay9 %></li>
		             <li class="group">Ten Week's Prior Stats</li>
		<li><% response.write "GLAZING: " & totalgy10 %></li>
		<li><% response.write "ASSEMBLY: " & totalay10 %></li>
-->		

            
        
<%

wcount=0
JFCHECKID=0


'THIS IS SLOW, BUT IT MAKES SURE THAT ALL WINDOWS RELATED TO A BATCH ARE COUNTED FROM MORE THEN TODAY'S LIST
'rs.filter = "ID > 0"
'rs.movefirst 
'
'do while not rs.eof
'rs2.filter = "ID > 0"
'
'
'		DIM Job, Floor, Dept, JFCHECKID, wcount
'		JOB = RS("JOB")
'		FLOOR = RS("JOB")
'		DEPT = RS("DEPT")
'		JFCHECKID = 0
'		
'		
'					do while not rs2.eof
'	
'
'					IF rs2("JOB") = RS("JOB") AND rs2("FLOOR") = rs("Floor") AND RS2("DEPT") = rs("DEPT") THEN
'					JFCHECKID = RS2("ID")
'					wcount = wcount + 1	
'					END IF
'				rs2.movenext
'				LOOP
'		
'		
'					IF JFCHECKID = "0" THEN
'					rs2.addnew 
'					rs2.fields("JOB") = RS("JOB")
'					rs2.fields("FLOOR") = RS("FLOOR")
'					rs2.fields("DEPT") = RS("DEPT")
'					rs2.fields("YEAR") = RS("YEAR")
'					rs2.fields("MONTH") = RS("MONTH")
'					rs2.fields("DAY") = RS("DAY")
'					rs2.fields("WEEK") = RS("WEEK")
'					rs2.fields("TAG") = 1
'					RS2.UPDATE
'					
'					ELSE	
'					
'					rs2.filter = "ID = " & JFCHECKID
'					rs2.fields("TAG") = rs2.Fields("TAG") + wcount
'					rs2.update
'					
'					end if
'				
'				wcount = 0
'				
'				
'	
'	
'			
'		 %>
'<% 
'	rs.movenext
'loop	
'	rs2.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
'	do while not rs2.eof
'	
'	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
'	if rs4.bof then
'	else
'	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
'	end if
'	rs2.movenext
'	
'	loop
'	
	
	%> 
 

        </ul>
        
        
      





</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing
rs4.close
set rs4=nothing
rs5.close
set rs5=nothing
'rs6.close
'set rs6=nothing

DBConnection.close
set DBConnection=nothing
%>

