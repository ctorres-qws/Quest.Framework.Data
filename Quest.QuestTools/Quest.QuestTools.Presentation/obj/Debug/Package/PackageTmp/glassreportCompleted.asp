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
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

    <style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT top 10000 * FROM Z_GLASSDB ORDER BY ID DESC"
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

       <form id="Complete" name="Complete"  method="GET" target="_self" selected="true" >  
        <ul id="Profiles" title=" Glass Report - All Completed" selected="true">
        <li>All Completed GLASS REPORT </li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	<li class='group'>Choose Completed Window items that have been SHIPPED 
	<input type='submit' value = 'Remove Selected Items' onClick="Complete.action='glassRemoveConf.asp?ticket=completed'; Complete.submit()"></li>
	<li class='group'>Choose Entire POs that have been SHIPPED 
	<input type='submit' value = 'Remove Entire PO' class = "blueButton" onClick="Complete.action='glassRemovePOConf.asp?ticket=completed'; Complete.submit()"></li>

<%
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>TimeLine</th><th>SHIP</th></tr>"
do while not rs.eof
	if isdate(RS("COMPLETEDDATE")) and Not isdate(RS("SHIPDATE")) then
		response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
		response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td>"
		response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=completed' target ='#_blank' >Time Line</a> </td>"
		response.write "<td><input type='checkbox' name='GID' value='" & RS("ID")& "'></td></tr>"
	end if
	rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
      </ul>
	  <input type = 'hidden' name = "ticket" value = "Completed" />	
		</form>

</body>
</html>
