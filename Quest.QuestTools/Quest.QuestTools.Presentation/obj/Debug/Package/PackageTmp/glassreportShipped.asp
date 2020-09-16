<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- Designed for Eric Bedeov, this report shows all items marked Shipped, Michael Bernholtz, November 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<style>
table{
zoom: 70%;
};
 </style>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("SELECT top 2000 * FROM Z_GLASSDB Where SHIPDATE <> NULL ORDER BY ID DESC",isSQLServer)
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

        <ul id="Profiles" title=" Glass Report - All Completed" selected="true">
        <li class='group'>All Units Marked Completed & Shipped  </li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
<%

response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>Notes</th><th>TimeLine</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td><td>" & RS("NOTES") & "</td>"
	response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=ship' target ='#_blank' >Time Line</a> </td>"
	response.write " </tr>"
	rs.movenext

loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

      </tbody></table></ul>

</body>
</html>
