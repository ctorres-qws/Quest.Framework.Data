<%
Response.Buffer = False
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Panel Reporting based on Awning Reporting to view information  -->
<!-- Michael Bernholtz, August 2016, Developed at Request of Lev Bedoev and Alex Goldbaum - Adapted mainly from Awning Report code-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Report</title>
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
    <div class="toolbar">
        <h1 id="pageTitle">Panels Report</h1>
        <a class="button leftButton" type="cancel" href="ScanPanelAll.asp" target='_self' >Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Awning Department" selected="true">
        
        


<li class="group"><a href="PanelReport.asp?RangeView=Today" target="_self" >View Today</a></li>
<li class="group"><a href="PanelReport.asp?RangeView=Week" target="_self" >View This Week</a></li>
<li class="group"><a href="PanelReport.asp?RangeView=Month" target="_self" >View This Month</a></li>
<li class="group"><a href="PanelReport.asp?RangeView=All" target="_self" >View All</a></li> 
<li ><a class="lightblueButton" href="ScanPanelAll.asp" target="_self" >Panel Scanner</a></li>
<li><table border='1' class = 'sortable' ><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Type</th><th>Employee</th><th>Department</th><th> Date</th></tr>

    <%
RangeView = Request.QueryString("RangeView")
'View ALL
strSQL = "Select * FROM X_BARCODEP ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"

if RangeView = "Today" then
strSQL = "Select * FROM X_BARCODEP WHERE DAY = " & DAY(NOW) & " AND MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Month" then
strSQL = "Select * FROM X_BARCODEP WHERE MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
if RangeView = "Week" then
strSQL = "Select * FROM X_BARCODEP WHERE WEEK = " & DatePart("ww", NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY DEPT ASC, JOB ASC, FLOOR ASC, TAG ASC"
end if
	
Cut = 0
Bend = 0
Ship = 0
Receive = 0	
	
Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & rs("JOB") & "</td>"
	response.write "<td>" & rs("FLOOR") & "</td>"
	response.write "<td>" & rs("Tag") & "</td>"
	response.write "<td>" & rs("Type") & "</td>"
	response.write "<td>" & rs("Employee") & "</td>"
	response.write "<td>" & rs("DEPT") & "</td>"
	
Select Case rs("DEPT")
	Case "Cut"
		Cut = Cut + 1
	Case "Bend"
		Bend = Bend + 1	
	Case "Ship"
		Ship = Ship + 1
	Case "Receive"
		Receive = Receive + 1
End Select
	
	response.write "<td>" & rs("DATETIME") & "</td>"
	response.write "</tr>"
	rs.movenext
loop
response.write "</table></li>"
%>
<li><b>Total</b> Cut: <%response.write Cut %>  Bend: <%response.write Bend %>   Ship: <%response.write Ship %>  Receive: <%response.write Receive %>  </li>
</ul>
<%
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
