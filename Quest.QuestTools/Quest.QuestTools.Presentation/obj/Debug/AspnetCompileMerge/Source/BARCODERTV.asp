<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- BARCODERTV.asp saved as BarocderTV-Database in July 2014-->
<!-- Barcode_V.asp maintains the database for this information and is run on a schedule-->
<!-- This is a new Main page that only reads from the tables every ten minutes. Should be less strain on the System-->

<!-- AUGUST 2014 - Started adding Drill Down Reports -->
<!-- DrillDown Report, ZipperRedReport, ZipperBlueReport-->

	<!--BARCODERTVEmail.aspx - is an E-mail report of this information in webmail form, it must be updated when the code here gets changed-->
	<!--February 2018 - yesterday hour changed to 7 to accomodate 6am night shift-->

  <title>Quest Dashboard</title>
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

<!--#include file="dbpath.asp"-->
<!--#include file="todayandyesterday.asp"-->

<% 
if chour < 7 then
	cDay = cYesterday
	cMonth = cMonthy
	cYear = cYeary
	Twodayago = CurrentDate -2
	cYesterday = Day(Twodayago)
	cMonthy = Month(TwoDayAgo)
	cYeary = Year(TwoDayAgo)
end if ' before seven am code (February 2018 change from 3am)

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





%>
</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle">Progress Today</h1>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">
<%
if rs.eof then
	response.write "<li> There is no information Available yet about today's activity </li>"
else
%>

 		<li class="group">Shipping Statistics</li>
		<%
		str_Yesterday = cYesterday & "-" & Left(MonthName(cMonthy),3) & "-" & cYeary 
		str_Today = cDay & "-" & Left(MonthName(cMonth),3) & "-" & cYear
		rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear
		response.write "<li>Report Period: " & str_Yesterday & " to " & str_Today & "</li>"
		response.write "<li>Windows Scanned to Trucks: " & rs("ShipScan") & "</li>"
		response.write "<li># of Trucks Opened: " & rs("TruckOpen") & "( " & rs("TruckOpenName") & ")</li>"
		response.write "<li># of Trucks Closed: " & rs("TruckClose") & "( " & rs("TruckCloseName") & ")</li>"
		%> 
		

		<li class="group">Dept Stats</li>
		<li><% response.write "Full Scans: " & RS("GlazingFull") & " - " & RS("SquareFoot") & "ft<sup>2</sup>" %></li>	
        <li><% response.write "Partial Scans: " & RS("GlazingPartial") %></li>
		<li><a href = "DrillDownReport.asp?Ticket=BarcoderTV&DEPT=ASSEMBLY" target="_self"><% response.write "ASSEMBLY: " & RS("ASSEMBLY") %></a></li>
		
		<li><a href = "PanelTV.asp?Ticket=BarcoderTV" target="_self"><% response.write "PANEL: " & RS("PANEL") %></a></li>

		<li><a href = "AwningReport.asp?RangeView=Today" target="_self"><% response.write "Awning Glaze: " & RS("AWNING") %></a></li>
		
		<li class="group">Glass Production</li>
        <li><a href = "GlassTV.asp?Ticket=BarcoderTV" target="_self"><% response.write "FOREL: " & RS("FOREL") %></a></li>	
        <li><a href = "GlassTV.asp?Ticket=BarcoderTV" target="_self"><% response.write "WILLIAN: " & RS("WILLIAN") %></a></li>	
		
		<li class="group">Zipper Machines</li>
        <li><a href = "zipperreport.asp?Ticket=BarcoderTV" target="_self"><% response.write "RED ZIPPER: " & RS("ZIPPERRED") %></a></li>	
        <li><a href = "zipperreport.asp?Ticket=BarcoderTV" target="_self"><% response.write "BLUE ZIPPER: " & RS("ZIPPERBLUE") %></a></li>	
		<!-- Old reports ZipperRedReport.asp and ZipperBlueReport.asp -->
		<% 
		if weekday(currentDate) = 2 then
			monday = 1
		else
			monday = 0
		end if
		
		if monday = 1 then
		%>
			<li class="group">Saturday's Stats</li>
		<%		
			cYesterday = Day(DateAdd("d",-2,Now))
			cMonthy = Month(DateAdd("d",-2,Now))
			cYeary =Year(DateAdd("d",-2,Now))
			cYesterday2 = Day(DateAdd("d",-3,Now))
			cMonthy2 = Month(DateAdd("d",-3,Now))
			cYeary2 =Year(DateAdd("d",-3,Now))
		
		else
		%>
			<li class="group">Yesterday's Glazing Statistics (<%= str_Yesterday %>)</li>
		
		<%
		end if
		
		rs.filter = ""
		rs.filter = "DAY = " & cYesterday & " AND Month = " & cMonthy & " AND YEAR = " & cYeary
		If Not rs.EOF Then
		%>
		<li><% response.write "Full Scans: " & RS("GlazingFull") & " - " & RS("SquareFoot") & "ft<sup>2</sup>" %></li>	
		<li><% response.write "Partial Scans: " & RS("GlazingPartial") %></li>
		<li><% response.write "ASSEMBLY: " & RS("ASSEMBLY") %></li>
		<li><% response.write "PANEL: " & RS("PANEL") %></li>
		<li><% response.write "AWNING: " & RS("AWNING") %></li>
		<li><% response.write "FOREL: " & RS("FOREL") %></li>
		<li><% response.write "WILLIAN: " & RS("WILLIAN") %></li>	
		<li><% response.write "RED ZIPPER: " & RS("ZIPPERRED") %></li>
		<li><% response.write "BLUE ZIPPER: " & RS("ZIPPERBLUE") %></li>
		<li class="group">Yesterday's Shipping Statistics (<%= str_Yesterday %>)</li>
		<%
		response.write "<li>Windows Scanned to Trucks: " & rs("ShipScan") & "</li>"
		response.write "<li># of Trucks Opened: " & rs("TruckOpen") & "( " & rs("TruckOpenName") & ")</li>"
		response.write "<li># of Trucks Closed: " & rs("TruckClose") & "( " & rs("TruckCloseName") & ")</li>"
		End If
		%> 

		<%
		if Monday = 1 then
		rs.filter = ""
		rs.filter = "DAY = " & cYesterday2 & " AND Month = " & cMonthy2 & " AND YEAR = " & cYeary2
		%>
		<li class="group">Friday's Glazing Statistics</li>
<li><% response.write "Full Scans: " & RS("GlazingFull") & " - " & RS("SquareFoot") & "ft<sup>2</sup>" %></li>	
        <li><% response.write "Partial Scans: " & RS("GlazingPartial") %></li>

		<li><% response.write "ASSEMBLY: " & RS("ASSEMBLY") %></li>
		
		<li><% response.write "PANEL: " & RS("PANEL") %></li>

		<li><% response.write "AWNING: " & RS("AWNING") %></li>
		
		<li class="group">Friday's Shipping Statistics</li>
		<%
		response.write "<li>Windows Scanned to Trucks: " & rs("ShipScan") & "</li>"
		response.write "<li># of Trucks Opened: " & rs("TruckOpen") & "( " & rs("TruckOpenName") & ")</li>"
		response.write "<li># of Trucks Closed: " & rs("TruckClose") & "( " & rs("TruckCloseName") & ")</li>"
		%> 
		
		
		<%
		ENd if
		%>
		 <li class="group">Assembly Activity</li>
        <%
        rs2.filter = "DEPT = 'ASSEMBLY'"
		do while not rs2.eof
		response.write "<li>" & rs2("JOB") & " " & rs2("FLOOR") & " (" & rs2("TODAY") & ") " & RS2("PROGRESS") & "/" & RS2("TOTAL") & "</li>"
		rs2.movenext
		loop
		%>
		 <li class="group">Glazing Activity</li>
        
		<%
		rs2.filter = "DEPT = 'GLAZING'"
		do while not rs2.eof
		response.write "<li>" & rs2("JOB") & " " & rs2("FLOOR") & " (" & rs2("TODAY") & ") " & RS2("PROGRESS") & "/" & RS2("TOTAL") & "</li>"
		rs2.movenext
		loop
		%>
		
		 <li class="group">CNC Machines</li>
		<%
		rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear
		response.write "<li><a href = 'DrillDownReportSaws.asp?Ticket=BarcoderTV' target='_self'>EW-JAMB " & rs("EWJAMB") & "</a></li>"
		response.write "<li><a href = 'DrillDownReportSaws.asp?Ticket=BarcoderTV' target='_self'>EW-WIDTH " & rs("EWWIDTH") & "</a></li>"
		response.write "<li><a href = 'DrillDownReportSaws.asp?Ticket=BarcoderTV' target='_self'>Q-JAMB " & rs("QJAMB") & "</a></li>"
		response.write "<li><a href = 'DrillDownReportSaws.asp?Ticket=BarcoderTV' target='_self'>Q-WIDTH " & rs("QWIDTH") & "</a></li>"
		%>
        </ul>
        
  
<% 
'Error Catching for No Record yet
end if

rs.close
set rs=nothing
rs2.close
Set rs2= nothing
DBConnection.close
set DBConnection = nothing

%>


</body>
</html>
