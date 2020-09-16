<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!--AllJobs Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>ALL jobs Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<%

	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=AllJobsReport.xls"
	Else
%>
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
 <!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

  <script type="text/javascript">
	$(document).ready( function () {
		$('#Job').DataTable({
			"iDisplayLength": 25
		});
	});
  
  </script>
<%
	End If
%>
    </head>
<body>

<%	
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_jobs ORDER BY JOB ASC"
	rs.Cursortype = GetDBCursorType
	rs.Locktype = GetDBLockType
	rs.Open strSQL, DBConnection
%>
<% If Request("Download") <> "YES" Then %>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>

        <ul id="Profiles" title=" Glass Report - All Active" selected="true">
        <li class='group'>All Jobs Summary </li>
        <li> <a class="whiteButton" href="ALLJOBSENTER.asp" target='_Self'>Add New PARENT Job</a></li>
		<li> <a class="lightblueButton" href="ALLJOBSCHILDENTER.asp" target='_Self'>Add New CHILD Job</a></li>
		<li><a class="greenButton" href="AllJobsReport.asp?Download=YES">Download Excel Copy</a></li>
         <div style='text-align: right; padding-right: 300px;'></a></div>
<% 
End If
response.write "<li><table border='1' class='Job' id ='Job'><thead><tr><th>Job</th><th>Full Name</th><th>Address</th><th>City</th><th>Country</th><th>EcoWall / Q4750</th><th># of Floors</th><th>Engineer</th><th>Manager</th><th>Parent</th><th>Import Tax ID</th><th>Job Status</th><th>Completed</th><th width ='150'>Edit</th></tr></thead><tbody>"

do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JOB") & "</td>"
	response.write "<td>" & RS("JOB_NAME") &"</td>"
	response.write "<td>" & RS("JOB_ADDRESS") & "</td>"
	response.write "<td>" & RS("JOB_CITY") & "</td>"
	response.write "<td>" & RS("JOB_COUNTRY") & "</td>"
	response.write "<td>" & RS("MATERIAL") & "</td>"
	response.write "<td>" & RS("FLOORS") & "</td>"
	response.write "<td>" & RS("Engineer") & "</td>"
	response.write "<td>" & RS("Manager") & "</td>"
	response.write "<td>" & RS("Parent") & "</td>"
	response.write "<td>" & RS("IMPORTERTAXID") & "</td>"
	response.write "<td>" & RS("JobStatus") & "</td>"
	response.write "<td>" & RS("Completed") & "</td>"
	If Request("Download") <> "YES" Then response.write "<td> <a class='lightblueButton' href='ALLJOBSEditForm.asp?JID=" & RS("ID") & "' target='_Self'>Edit this Job</a></td>"
	response.write " </tr>"

	rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing



%>
	</tbody></table>
      </ul>                 
            
     
               
</body>
</html>
