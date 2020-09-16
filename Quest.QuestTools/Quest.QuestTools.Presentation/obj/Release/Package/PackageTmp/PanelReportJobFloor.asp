<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Panel Reporting for specific Job and Floor  -->
<!-- Michael Bernholtz, August 2016, Takes each tag and checks for each type of scan.-->


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
	<% Server.ScriptTimeout = 300 %> 

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Panels Report</h1>
        <a class="button leftButton" type="cancel" href="ScanPanelAll.asp" target='_self' >Reports</a>
        </div>
   

        <ul id="Profiles" title="Glass Report - Panel Department" selected="true">


<li><form id="screen1" title="Panels" class="panel" name="PanelForm" action="PanelReportJobFloor.asp" method="GET" selected="true">
<fieldset>
   <div class="row">   
            <label>Job </label>
            <input type="text" name='Job' id='Job' value = "<%response.write Job%>" >
		</div>
	<div class="row">     
            <label>Floor </label>
            <input type="text" name='Floor' id='Floor' value = "<%response.write Floor%>">
		</div>
		<a class="whiteButton" onClick=" PanelForm.submit()">View Panels </a><BR>
</fieldset>
</form> 
</li>

<%
Job = Request.QueryString("Job")
FLoor = Request.Querystring("Floor")
if Job = "" and FLoor = "" then
strSQL = "Select top 200 * FROM X_BARCODEP ORDER BY TAG ASC, DEPT ASC"
JobFloor = "ALL"
else
	if (Job <> "" and FLoor = "") or (Job <> "" and FLOOR = "ALL") then
		strSQL = "Select * FROM X_BARCODEP WHERE JOB = '" & Job & "' ORDER BY FLOOR ASC, TAG ASC, DEPT ASC"
		JobFloor = Job & " - ALL FLOORS"
	else
		strSQL = "Select * FROM X_BARCODEP WHERE JOB = '" & Job & "' AND FLOOR = '" & Floor & "' ORDER BY TAG ASC, DEPT ASC"
		JobFloor = Job & " " & FLoor
	end if
end if
%>
<li><h2>Showing Data for <%Response.write JobFloor %></h2></li>
<li><table border='1' class = 'sortable' ><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Type</th><th>Employee</th><th>Department</th><th> Date</th></tr>

    <%

	
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
