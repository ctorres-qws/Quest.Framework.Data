<%Response.Buffer = false%>                        
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

	<!-- Updated May 2019 to include Texas Functionality - Ariel Aziza Michael Bernholtz-->
<!-- Scanned Door Report - Showing all Door items as scanned -->
<!-- Originally based on Awnign Report Designed Jan 2020-->



<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Door Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
	<!--#include file="dbpath.asp"-->
	<%
	ScanMode = TRUE
	%>
	<!--#include file="Texas_dbpath.asp"-->
<div class="toolbar">
        <h1 id="pageTitle">Door Produced</h1>
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
   
   
        <ul id="Profiles" title="Glass Report - Door Department" selected="true">
        
        


<li class="group"><a href="DoorReport.asp?RangeView=Today" target="_self" >View Today</a></li>
<li class="group"><a href="DoorReport.asp?RangeView=Week" target="_self" >View This Week</a></li>
<li class="group"><a href="DoorReport.asp?RangeView=Month" target="_self" >View This Month</a></li>
<li class="group"><a href="DoorReport.asp?RangeView=All" target="_self" >View All(Year)</a></li> 
<li><table border='1' class = 'sortable' ><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Type</th><th>Employee</th><th>Department</th><th> Date</th></tr>

    <%
RangeView = Request.QueryString("RangeView")
'View Today
strSQL = "Select * FROM X_BARCODESW WHERE DAY = " & DAY(NOW) & " AND MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"

if RangeView = "All" then
strSQL = "Select * FROM X_BARCODESW WHERE YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Month" then
strSQL = "Select * FROM X_BARCODESW WHERE MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Week" then
strSQL = "Select * FROM X_BARCODESW WHERE WEEK = " & DatePart("ww", NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
	
FrameSL = 0
GlassSL = 0
SWDoor = 0
SunviewDoor = 0	
	
Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType

if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else
	rs.Open strSQL, DBConnection
end if


do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & rs("JOB") & "</td>"
	response.write "<td>" & rs("FLOOR") & "</td>"
	response.write "<td>" & rs("Tag") & "</td>"
	response.write "<td>" & rs("Type") & "</td>"
	response.write "<td>" & rs("Employee") & "</td>"
	response.write "<td>" & rs("DEPT") & "</td>"
	
Select Case rs("DEPT")
	Case "SlidingFrame"
		FrameSL = FrameSL + 1
	Case "SlidingGlass"
		GlassSL = GlassSL + 1	
	Case "SwingDoor"
		SWDoor = SWDoor + 1
	Case "SunviewDoor"
		SunviewDoor = SunviewDoor + 1
End Select
	
	response.write "<td>" & rs("DATETIME") & "</td>"
	response.write "</tr>"
	rs.movenext
loop
response.write "</table></li>"
%>
<li><b>Total</b> SL Frame: <%response.write FrameSL %> SL Glass: <%response.write GlassSL %>  SW Door: <%response.write SWDoor %>  Sunview Door: <%response.write SunviewDoor %>  </li>
</ul>
<%
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
        
               
</body>
</html>
