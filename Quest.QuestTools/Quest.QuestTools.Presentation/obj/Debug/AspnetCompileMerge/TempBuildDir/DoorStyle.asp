<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->
<!-- Date: June 14, 2019
	 Modified By: Michelle Dungo
	 Changes: Dynamic drop down of job names which only shows parent jobs
			  Add script to refresh parameters when job dropdown value changes	 
-->
<%
' Generate list from Z_Jobs in Quest DB where Job has not been completed using Parent field or Job field if Parent field is empty
Set DBConnectionJob = Server.CreateObject("ADODB.Connection")
Set rsJob = Server.CreateObject("ADODB.Recordset")
DSN = GetConnectionStr(False) ' connect to access for job list
DBConnectionJob.Open DSN

strSQL = "SELECT DISTINCT Parent "
strSQL = strSQL & "FROM Z_Jobs Where Parent <> '' and Completed = False "
'strSQL = strSQL & " and Parent NOT LIKE 'AA%' " 'exclude test job
'start: include jobs with no Parent field - commented out, prefer fix to be done at job table
'strSQL = strSQL & "UNION SELECT DISTINCT Job as Parent " 
'strSQL = strSQL & "FROM Z_Jobs Where Parent = '' and Completed = False "
'strSQL = strSQL & " and Job NOT LIKE 'AA%' " 'exclude test job
'end: include jobs with no Parent field
rsJob.Cursortype = 2
rsJob.Locktype = 3
rsJob.Open strSQL, DBConnectionJob

While Not rsJob.EOF
	ParentJob = ParentJob & "<option value=" & rsJob("Parent") & ">" & rsJob("Parent") & "</option>"
rsJob.MoveNext
Wend

rsJob.Close
DBConnectionJob.Close
set rsJob = nothing
set DBConnectionJob = nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Door Styles</title>
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
  <script>
	function periodChange() {
		enter.action = "DoorStyleEnter.asp"
		enter.qwsAction.value = 'JOB_CHANGE';
		enter.submit();
	}
  </script>	  

  <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
  <script src="sorttable.js"></script>
</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
            <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job</a>
    </div>
    
    <form id="edit" title="Select Door Style by Job" class="panel" name="edit" action="DoorStylebyJob.asp" method="GET" target="_self" selected="true" > 
    <h2>Door Styles by Job</h2>

	<fieldset>
		<div class="row">
			<label> Job </label>
				<Select name='Parent' id='Parent'>
					<option value="AllJobs">All Jobs</option>
					<%=ParentJob%>
				</Select>	
		</div>
	</fieldset>

	<br>
	<a class="whiteButton" href="javascript:edit.submit()">Search ALL Door Styles by Job </a>
	<a class="whiteButton" href="DoorStyleEnter.asp">Add New Style </a>
	<br><br>
	<ul id="Profiles" title="Door Styles - Job Search" selected="true">
	<%
	gi_Mode = c_MODE_SQL_SERVER
	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select	

	Function Process(isSQLServer)

		DBOpen DBConnection, isSQLServer

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM StylesDoor ORDER BY Job, Name ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection		

		response.write "<li class='group'>All Door Styles</li>"
		response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
		response.write "<li><table border='1' class='sortable'><tr><th>Name</th><th>Job</th><th>Interior Door Type</th><th>Exterior Door Type</th></tr>"
		do while not rs.eof
			response.write "<tr>"
			response.write "<td>" & rs("Name") & "</td>"
			response.write "<td>" & rs("Job") &"</td>"
			response.write "<td>" & rs("IntDoorType") & "</td>"
			response.write "<td>" & rs("ExtDoorType") &"</td>"			
			response.write "<td><a href =><a href='DoorStyleEditForm.asp?cid=" & rs("ID") & "' target='_self' >Manage</td>"
			response.write " </tr>"
			rs.movenext
		loop
		response.write "</table></li>"

		rs.close
		set rs = nothing
		DBConnection.close 
		set DBConnection = nothing
	End Function
	%>		
	</form> 
</body>
</html>


