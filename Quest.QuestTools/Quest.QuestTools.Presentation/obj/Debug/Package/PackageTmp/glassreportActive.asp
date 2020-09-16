<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
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
strSQL = "SELECT * FROM Z_GLASSDB WHERE [HIDE] IS NULL ORDER BY ID DESC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

        <ul id="Profiles" title=" Glass Report - All Active" selected="true">
<% 
response.write "<li class='group'>All Active GLASS REPORT </li>"
response.write "<li class='group'>" & Request.ServerVariables("REMOTE_ADDR") & "</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>Status</th><th>Notes</th><th>TimeLine</th></tr>"
do while not rs.eof
	if not isdate(RS("COMPLETEDDATE")) then
		response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
		response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td>"
%>
			<!--#include file="GlassStatus.inc"-->
<%
		Response.write "<td>" & Status & "</td> " ' NStatus
		response.write "<td>"
		if rs("NOTES") <> "" then
			response.write RS("NOTES") & " - "
		end if
		response.write "<a href='glassOptimaNote.asp?gid="  & RS("ID") & "&ticket=active' target ='#_Self' > Add</a> </td>"
		response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=active' target ='#_blank' >Time Line</a> </td>"
		response.write " </tr>"
	end if
	rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
      </ul>
</body>
</html>
