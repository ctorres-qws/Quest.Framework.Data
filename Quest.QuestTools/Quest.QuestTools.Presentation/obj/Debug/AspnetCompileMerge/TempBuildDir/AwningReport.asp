<%Response.Buffer = false%>                        
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

	
<!-- Scanned Awning Report - Showing all Awning items as scanned -->
<!-- Originally based on Glazing and Backorder code 2015-->
<!-- Updated May 2019 to include Texas Functionality - Ariel Aziza Michael Bernholtz-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Awning Report</title>
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
        <h1 id="pageTitle">Awnings Produced</h1>
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
   
   
        <ul id="Profiles" title="Glass Report - Awning Department" selected="true">
        
        


<li class="group"><a href="AwningReport.asp?RangeView=Today" target="_self" >View Today</a></li>
<li class="group"><a href="AwningReport.asp?RangeView=Week" target="_self" >View This Week</a></li>
<li class="group"><a href="AwningReport.asp?RangeView=Month" target="_self" >View This Month</a></li>
<li class="group"><a href="AwningReport.asp?RangeView=All" target="_self" >View All(Year)</a></li> 
<!--<li ><a class="lightblueButton" href="ScanAwning.asp" target="_self" >Awning Scanner</a></li>-->
<li><table border='1' class = 'sortable' ><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Type</th><th>Employee</th><th>Department</th><th> Date</th></tr>

    <%
RangeView = Request.QueryString("RangeView")
'View Today
strSQL = "Select * FROM X_BARCODEOV WHERE DAY = " & DAY(NOW) & " AND MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"

if RangeView = "All" then
strSQL = "Select * FROM X_BARCODEOV WHERE YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Month" then
strSQL = "Select * FROM X_BARCODEOV WHERE MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Week" then
strSQL = "Select * FROM X_BARCODEOV WHERE WEEK = " & DatePart("ww", NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
	
Frame = 0
Sash = 0
Glaze = 0
Mount = 0	
	
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
	Case "FrameAssemble"
		Frame = Frame + 1
	Case "SashAssemble"
		Sash = Sash + 1	
	Case "SashGlaze"
		Glaze = Glaze + 1
	Case "WindowMount"
		Mount = Mount + 1
End Select
	
	response.write "<td>" & rs("DATETIME") & "</td>"
	response.write "</tr>"
	rs.movenext
loop
response.write "</table></li>"
%>
<li><b>Total</b> Frame: <%response.write Frame %>  Sash: <%response.write Sash %>   Glaze: <%response.write Glaze %>  Mount: <%response.write Mount %>  </li>
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
