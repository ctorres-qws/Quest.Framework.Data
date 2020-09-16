
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<!--Based on BarcodeWeekly - New Glazing Tool counts based on X_GLAZING	-->
	<!-- Assembly still based on X_Barcode -->
	<!--Updated May 2019 to include Texas Database, Ariel Aziza Michael Bernholtz -->

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
	<!--#include file="dbpath.asp"-->
	<%
	ScanMode = TRUE
	%>
	<!--#include file="Texas_dbpath.asp"-->
<!--#include file="todayandyesterday.asp"-->
<%

weekNumber = DatePart("ww", currentDate)
lastweek  = DatePart("ww", DateAdd("d", -7, currentDate))
lastweek2 = DatePart("ww", DateAdd("d", -14, currentDate))

cyeary1 = cyeary
cyeary2 = cyeary

if Weeknumber = 1 then
	cyeary1 = cyeary-1
	cyeary2 = cyeary-1
end if
if Weeknumber = 2 then
	cyeary1 = cyeary
	cyeary2 = cyeary-1 
	
end if

Set rs = Server.CreateObject("adodb.recordset")
SQL = "Select * FROM X_BARCODE WHERE DEPT = 'ASSEMBLY' AND (YEAR = " & CYEAR & " OR YEAR = " & CYEARY2 & ") AND (WEEK = " & lastweek2 & " OR WEEK = " & lastweek & " OR WEEK = " & Weeknumber & ") ORDER BY DATETIME DESC"
'rs.Cursortype = 1
'rs.Locktype = 3
'rs.Open SQL, DBConnection

if CountryLocation = "USA" then
	Set rs = GetDisconnectedRS(SQL, DBConnection_Texas)
else
	Set rs = GetDisconnectedRS(SQL, DBConnection)
end if



totala = 0
totalay = 0
totalay2= 0


' This week
rs.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear

	totala = rs.recordcount
	
'Last Week	
if not rs.eof then
rs.movefirst
end if
rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYeary1
	totalay = rs.RecordCount

'two week's prior stats
if not rs.eof then
rs.movefirst
end if
rs.filter = "WEEK = " & lastweek2 & " AND YEAR = " & cYeary2
totalay2 = rs.RecordCount

rs.close
set rs=nothing

Set rs2 = Server.CreateObject("adodb.recordset")
SQL2 = "Select * FROM X_GLAZING ORDER BY DATETIME DESC"
'rs2.Cursortype = 0 '2
'rs2.Locktype = 1   '3
'rs2.Open SQL2, DBConnection

if CountryLocation = "USA" then
	Set rs2 = GetDisconnectedRS(SQL2, DBConnection_Texas)
else
	Set rs2 = GetDisconnectedRS(SQL2, DBConnection)
end if



totalgf = 0
totalgfy = 0
totalgfy2= 0
totalgp2 = 0
totalgpy = 0
totalgpy2= 0

' This week
rs2.filter = "WEEK = " & weekNumber & " AND YEAR = " & cYear
Do while not rs2.eof
IF rs2("FIRSTCOMPLETE") = "TRUE" then
totalgf = totalgf + 1
else
totalgp = totalgp + 1
end if
rs2.movenext
loop
	
'Last Week	
rs2.filter = ""
rs2.filter = "WEEK = " & lastweek & " AND YEAR = " & cYeary1
Do while not rs2.eof
IF rs2("FIRSTCOMPLETE") = "TRUE" then
totalgfy = totalgfy + 1
else
totalgpy = totalgpy + 1
end if
rs2.movenext
loop

'two week's prior stats
rs2.filter = ""
rs2.filter = "WEEK = " & lastweek2 & " AND YEAR = " & cYeary2
Do while not rs2.eof
IF rs2("FIRSTCOMPLETE") = "TRUE" then
totalgfy2 = totalgfy2 + 1
else
totalgpy2 = totalgpy2 + 1
end if
rs2.movenext
loop

rs2.close
set rs2 = nothing


DBConnection.close
set DBConnection=nothing

DBConnection_Texas.close
set DBConnection_Texas=nothing
%>

</head>
<body onload="startTime()" >

<div class="toolbar">
        <h1 id="pageTitle">Assembly/Glazing</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "indexTexas.html#_Report"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Report"
				HomeSiteSuffix = ""
			end if 	
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
<ul id="screen1" title="Quest Dashboard" selected="true">

		<li class="group">This Week's Stats <% response.write CountryLocation %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "Full Glaze: " & totalgf %></li>
		<li><% response.write "Partial Glaze: " & totalgp %></li>
		
        	<li class="group">Last Week's Stats</li>
		<li><% response.write "ASSEMBLY: " & totalay %></li>
		<li><% response.write "Full Glaze: " & totalgfy %></li>
		<li><% response.write "Partial Glaze: " & totalgpy %></li>
             <li class="group">Two Week's Prior Stats</li>
		<li><% response.write "ASSEMBLY: " & totalay2 %></li>
		<li><% response.write "Full Glaze: " & totalgfy2 %></li>
		<li><% response.write "Partial Glaze: " & totalgpy2 %></li>

    </ul>

</body>
</html>
