<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			
			<!--BARCODERTV-email.aspx - is an E-mail report of BARCODERTV information in webmail form, it must be updated when the code here gets changed-->
			<!--February 2018 - yesterday hour changed to 7 to accomodate 6am night shift-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>

 

<!--#include file="todayandyesterday.asp"-->
<% 
	currentDate = Now
	if Hour(currentDate) < 7 then
		cDay = cYesterday
		cMonth = cMonthy
		cYear = cYeary
	end if
	
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM V_REPORT1 ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM V_REPORT2"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection	

if weekday(currentDate) = 2 then
	Monday = 1
	
	if Hour(Now) < 7 then
		cYesterday = Day(DateAdd("d",-3,currentDate))
		cMonthy = Month(DateAdd("d",-3,currentDate))
		cYeary =Year(DateAdd("d",-3,currentDate))
	else
		cYesterday = Day(DateAdd("d",-2,currentDate))
		cMonthy = Month(DateAdd("d",-2,currentDate))
		cYeary =Year(DateAdd("d",-2,currentDate))	
	end if
else 

	Monday = 0
	if Hour(Now) < 7 then
		cYesterday = Day(DateAdd("d",-2,currentDate))
		cMonthy = Month(DateAdd("d",-2,currentDate))
		cYeary =Year(DateAdd("d",-2,currentDate))
	else
		cYesterday = Day(DateAdd("d",-1,currentDate))
		cMonthy = Month(DateAdd("d",-1,currentDate))
		cYeary =Year(DateAdd("d",-1,currentDate))
	end if
end if





%>
</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

<% 
		str_Yesterday = cYesterday & "-" & Left(MonthName(cMonthy),3) & "-" & cYeary 
		str_Today = cDay & "-" & Left(MonthName(cMonth),3) & "-" & cYear
		response.write "<b>Report Period: " & str_Yesterday & " to " & str_Today & "</b><br /><br />" 
%>

		<b><u>Shipping Statistics</u></b>
		
		<%
		response.write "<li>Windows Scanned to Trucks: " & rs("ShipScan") & "</li>"
		response.write "<li># of Trucks Opened: " & rs("TruckOpen") & "( " & rs("TruckOpenName") & ")</li>"
		response.write "<li># of Trucks Closed: " & rs("TruckClose") & "( " & rs("TruckCloseName") & ")</li>"
		%> 
	
		<br>

		<b><u>Glass Production Stats</u></b>
		<br>
		<li><% response.write "GLAZING FULL: " & RS("GLAZINGFull") &"  - " & RS("SquareFoot") & "ft<sup>2</sup>" %></li>
        <li><% response.write "GLAZING PARTIAL: " & RS("GLAZINGPartial") %></li>
		<li><% response.write "ASSEMBLY: " & RS("ASSEMBLY") %></li>
		<li><% response.write "PANEL: " & RS("PANEL") %></li>
        <li><% response.write "Awning Glaze: " & RS("AWNING") %></li>
		<br>
			<b><u>Glassline Stats</u></b>
        <li><% response.write "Forel Glass: " & RS("FOREL") %></li>
		<li><% response.write "Willian Glass: " & RS("WILLIAN") %></li>
			<b><u>Zipper Stats</u></b>
		<li><% response.write "Red Zipper: " & RS("ZIPPERRED") %></li>
		<li><% response.write "Blue Zipper: " & RS("ZIPPERBLUE") %></li>
		<br>

<%
		if Monday = 1 then
%>
			<br>
			<b><u>Saturday's Production Stats</u></b>
<%				
		else
%>
			<br>
			<b><u>Yesterday's Production Stats (<%= str_Yesterday %>)</u></b>
<%			
		end if
rs.filter = ""
rs.filter = "DAY = " & cYesterday & " AND Month = " & cMonthy & " AND YEAR = " & cYeary	
%>
		
		
		<li><% response.write "Glazing Full: " & RS("GLAZINGFULL") &"  - " & RS("SquareFoot") & "ft<sup>2</sup>" %></li>
        <li><% response.write "Glazing Partial: " & RS("GLAZINGPARTIAL") %></li>
		<li><% response.write "ASSEMBLY: " & RS("ASSEMBLY") %></li>
		<li><% response.write "PANEL: " & RS("PANEL") %></li>
        <li><% response.write "AWNING: " & RS("AWNING") %></li>
		<li><% response.write "Forel Glass: " & RS("FOREL") %></li>
		<li><% response.write "Willian Glass: " & RS("WILLIAN") %></li>
		<li><% response.write "Red Zipper: " & RS("ZIPPERRED") %></li>
		<li><% response.write "Blue Zipper: " & RS("ZIPPERBLUE") %></li>
<%
		if Monday = 1 then
%>
			<br>
			<b><u>Saturday's Shipping Stats</u></b>
<%				
		else
%>
			<br>
			<b><u>Yesterday's Shipping Stats (<%= str_Yesterday %>)</u></b>
<%			
		end if
%>
		<li>Windows Scanned to Truck: <% response.write  RS("Shipscan") %></li>
		<li># of Trucks Opened:	<% response.write RS("TruckOpen") & "( " & rs("TruckOpenName") %>)</li>
        <li># of Trucks Closed: <% response.write RS("TruckClose") & "( " & rs("TruckCloseName") %>)</li>
			
		
<%
rs.filter = ""
rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear		
%>
		<br>
       <b><u>Assembly Activity Today </u></b>
<%
        rs2.filter = "DEPT = 'ASSEMBLY'"
		do while not rs2.eof
		response.write "<li>" & rs2("JOB") & " " & rs2("FLOOR") & " (" & rs2("TODAY") & ") " & RS2("PROGRESS") & "/" & RS2("TOTAL") & "</li>"
		rs2.movenext
		loop
%>
		<br>
	   <b><u>Glazing Activity Today </u></b>
<%
		rs2.filter = "DEPT = 'GLAZING'"
		do while not rs2.eof
		response.write "<li>" & rs2("JOB") & " " & rs2("FLOOR") & " (" & rs2("TODAY") & ") " & RS2("PROGRESS") & "/" & RS2("TOTAL") & "</li>"
		rs2.movenext
		loop
%>
	
	<br>

		<b><u>CNC Machines</u></b>
		
<%
		response.write "<li>EW-JAMB " & rs("EWJAMB") & "</li>"
		response.write "<li>EW-WIDTH " & rs("EWWIDTH") & "</li>"
		response.write "<li>Q-JAMB " & rs("QJAMB") & "</li>"
		response.write "<li>Q-WIDTH " & rs("QWIDTH") & "</li>"
	%> 

	
        </ul>
  
<% 

rs.close
set rs=nothing
rs2.CLOSE
Set rs2= nothing
DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
