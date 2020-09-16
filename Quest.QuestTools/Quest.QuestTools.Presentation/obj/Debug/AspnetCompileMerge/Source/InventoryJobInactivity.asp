<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Inventory Job Inactivity</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	<script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
DBConnection.Close
DBConnection.Open GetConnectionStr(true)

		oldDate = DateAdd("m",-3,Date())
	
Set rs = Server.CreateObject("adodb.recordset")

strSQL = "SELECT job, max(modifydate) LastActivity, min(modifydate) FirstActivity, DateDiff(m, max(modifydate), getdate()) as Diff, DateDiff(m, min(modifydate), getdate()) as mDiff FROM "
strSQL = strSQL & "(select case when len(Replace(Replace(Colour,'Int.',''),'Ext.','')) > 3 then Replace(Replace(JobComplete,'Int.',''),'Ext.','') else Replace(Replace(Colour,'Int.',''),'Ext.','') end as Job, B.* from y_invlog B WHERE Warehouse IN ('WINDOW PRODUCTION')) b "
strSQL = strSQL & "where len(job) = 3 and job <> 'AAA' group by job order by job "


'strSQL = FixSQL("SELECT * FROM Y_INV WHERE (DATEIN < #" & oldDate & "# OR DATEIN = NULL) AND Warehouse = 'NASHUA' ORDER BY PART ASC, DATEIN DESC ")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory Old Jobs</a>
        </div>

        <ul id="Profiles" title="Inventory Job Inactivity" selected="true">

<% 

response.write "<li class='group'>Inventory</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Job</th><th>Last Activity</th><th>Months of Inactivity</th><th>First Activity</th><th>Started Months Ago</th></tr>"



do while not rs.eof

'If rs("Diff") >= 6 or rs("mDiff") >= 10 Then
If rs("Diff") >= 6 or (rs("mDiff") >= 10 and rs("Diff") > 1) Then

Response.write "<tr>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td align='center'>" & rs("LastActivity") & " </td>"
Response.write "<td align='right'>" & rs("Diff") & " </td>"
Response.write "<td align='center'>" & rs("FirstActivity") & " </td>"
Response.write "<td align='right'>" & rs("mDiff") & " </td>"
Response.write "</tr>"

End If

rs.movenext
loop
Response.write "</table></li>"

rs.close
set rs = nothing

DBConnection.close
Set DBConnection = nothing

%>
      <li></li>
	  </ul>
</body>
</html>
