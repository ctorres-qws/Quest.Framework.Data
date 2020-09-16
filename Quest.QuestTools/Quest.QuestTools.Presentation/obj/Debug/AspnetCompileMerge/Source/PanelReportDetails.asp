                       
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
        <a class="button leftButton" type="cancel" href="index.html#_Report" target='_self' >Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Awning Department" selected="true">
        
        


<li class="group"><a href="PanelReportDetails.asp?RangeView=Today" target="_self" >View Today</a></li>
<li class="group"><a href="PanelReportDetails.asp?RangeView=Week" target="_self" >View This Week</a></li>
<li class="group"><a href="PanelReportDetails.asp?RangeView=Month" target="_self" >View This Month</a></li>
<li ><a class="lightblueButton" href="ScanAwning.asp" target="_self" >Panel Scanner</a></li>
<li><table border='1' class = 'sortable' ><tr><th>Job / FLoor</th><th>Cut</th><th>Bend</th><th>Ship</th><th>Receive</th></tr>

    <%
RangeView = Request.QueryString("RangeView")
if RangeView = "" then
RangeView = "Today"
end if

if RangeView = "Month" then
strSQL = "Select * FROM X_BARCODEP WHERE MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY JOB ASC, FLOOR ASC, DEPT ASC, TAG ASC"
end if
if RangeView = "Week" then
strSQL = "Select * FROM X_BARCODEP WHERE WEEK = " & DatePart("ww", NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY JOB ASC, FLOOR ASC, DEPT ASC, TAG ASC"
end if
if RangeView = "Today" then
strSQL = "Select * FROM X_BARCODEP WHERE DAY = " & DAY(NOW) & " AND MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY JOB ASC, FLOOR ASC, DEPT ASC, TAG ASC"
end if
	
	
	
Job1 = ""
Job2 = ""
Cut = 0
Bend = 0
Ship = 0
Receive = 0	
TotalCut = 0
TotalBend = 0
TotalShip = 0
TotalReceive = 0	
	
Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
Job1 = rs("JOB") & RS("FLOOR")
Job2 = rs("JOB") & RS("FLOOR")
do while not rs.eof
	Job2 = Job1
	Job1 = rs("JOB") & RS("FLOOR")
	Select Case rs("DEPT")
			Case "Cut"
				Cut = Cut + 1
				TotalCut = TotalCut + 1
			Case "Bend"
				Bend = Bend + 1
				TotalBend = TotalBend + 1				
			Case "Ship"
				Ship = Ship + 1
				TotalShip = TotalShip + 1
			Case "Receive"
				Receive = Receive + 1
				TotalReceive = TotalReceive + 1
		End Select
	if Job1 = Job2 then	
	else
	response.write "<tr>"
	response.write "<td>" & rs("JOB") & "</td>"
	response.write "<td>" & cut & "</td>"
	response.write "<td>" & bend & "</td>"
	response.write "<td>" & ship & "</td>"
	response.write "<td>" & receive & "</td>"
	response.write "</tr>"
	
	Cut = 0
	Bend = 0
	Ship = 0
	Receive = 0	
	end if
	
	
rs.movenext
loop
	response.write "<tr>"
	response.write "<td>" & Job1 & "</td>"
	response.write "<td>" & cut & "</td>"
	response.write "<td>" & bend & "</td>"
	response.write "<td>" & ship & "</td>"
	response.write "<td>" & receive & "</td>"
	response.write "</tr>"
	

response.write "</table></li>"
%>
<li><b>Total</b> Cut: <%response.write TotalCut %>  Bend: <%response.write TotalBend %>   Ship: <%response.write TotalShip %>  Receive: <%response.write TotalReceive %>  </li>
</ul>
<%
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
