<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>All Jobs Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_Jobs ORDER BY JOB ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Jobs/Colour</a>
    </div>

<ul id="Profiles" title="Global Variables" selected="true">
	<li class='group'>All Jobs - Global Variables</li>
	<li><table border='1' class='sortable'><Thead>
		<tr>
			<th>ID</th><th>Job</th><th>Job Name</th><th>Parent</th>
			<th>Material</th><th>Frame Style</th><th>Sill ?</th>
			<th>Exterior colour</th><th>Interior Colour</th><th>Manager</th>
			<th>Manager Email</th><th>Completed</th>
		</tr></thead><tbody>

<%
do while not rs.eof
	
		response.write "<tr>"
		response.write "<td>" & RS("ID") & "</td>"
		response.write "<td>" & RS("JOB") & "</td>"
		response.write "<td>" & RS("JOB_Name") & "</td>"
		response.write "<td>" & RS("PARENT") & "</td>"
		response.write "<td>" & RS("MATERIAL") & "</td>"
		response.write "<td>" & RS("FRAMESTYLE") & "</td>"
		response.write "<td>" & RS("SILL") & "</td>"
		response.write "<td>" & RS("EXT_Colour") & "</td>"
		response.write "<td>" & RS("INT_Colour") & "</td>"
		response.write "<td>" & RS("Manager") & "</td>"
		response.write "<td>" & RS("ManagerEmail") & "</td>"
		response.write "<td>" & RS("Completed") & "</td>"	
		response.write " </tr>"
	rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
      </Tbody></Table></li></ul>
</body>
</html>
