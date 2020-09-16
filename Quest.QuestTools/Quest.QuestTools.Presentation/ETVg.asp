<!--#include file="dbpath.asp"-->
	<!--Updated May 2014 to prevent timeout-->    
		<!-- Updated May 2019 - to Recognize Employee VS Line Number as some have same code -->
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

  <% 

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

'Call SetTestDate(cDay, cMonth, cYear)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_GLAZING WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'GLAZING' order by FIRSTCOMPLETE DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.CacheSize = 100
rs.Open strSQL, DBConnection
	
Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT xE.* From X_EMPLOYEES xE WHERE Number IN(SELECT Employee FROM X_GLAZING WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'GLAZING') And Line = 'Glazing'"
rs5.Cursortype = GetDBCursorType
rs5.Locktype = GetDBLockType
rs5.CacheSize = 100
rs5.Open strSQL5, DBConnection

totalg = 0

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Glazing by Employee" selected="true">
    
    <li class="group">Today's Stats</li>
	<li class="group">Blue = Full Glaze / Red = Partial Glaze</li>
    
    <%

DO WHILE NOT RS5.EOF

rs.Close
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_GLAZING WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'GLAZING' And Employee='" & RS5("NUMBER") & "' order by FIRSTCOMPLETE DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

'rs.movefirst
TOTALG = 0
TOTALG2 = 0
Do while not rs.eof

IF rs("EMPLOYEE") = RS5("NUMBER") then
	IF UCASE(rs("FirstComplete")) = "TRUE" then
		totalG = totalG + 1
		IF RS5("SHIFT") = "0" then
			response.write "<img src='bluesquare.gif'>"
		ELSE
			response.write "<img src='dkbluesquare.gif'>"
		END IF
	ELSE
		totalG2 = totalG2 + 1
		IF RS5("SHIFT") = "0" then
			response.write "<img src='redsquare.gif'>"
		ELSE
			response.write "<img src='dkredsquare.gif'>"
		END IF
	END IF
END IF


rs.movenext
loop

IF (TOTALG = 0) AND (TOTALG2 = 0) THEN 
ELSE
RESPONSE.WRITE "<li>"
%><a href="ETVadet.asp?ticket=TodayGlazing&employee=<% RESPONSE.WRITE RS5("NUMBER") %>&first=<% response.write rs5("first") %>&last=<% response.write rs5("last") %>" target="_Self"><% RESPONSE.WRITE RS5("NUMBER") & " " & RS5("LAST") & ", " & RS5("FIRST") & " : " & totaLG & ", " & totalG2 %></a></li><%
END IF 

RS5.MOVENEXT
LOOP

RESPONSE.WRITE "</UL>"

rs.close
set rs=nothing
rs5.close
set rs5=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>