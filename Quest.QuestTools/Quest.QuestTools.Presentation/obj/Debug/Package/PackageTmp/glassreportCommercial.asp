<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Commercial Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

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
strSQL = "SELECT * FROM Z_GLASSDB WHERE SHIPDATE IS NULL AND [DEPARTMENT] = 'Commercial' ORDER BY ID ASC"
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

		<form id="Complete" name="Complete"  method="GET" target="_self" selected="true" >  
        <ul id="Profiles" title="Glass Report - Commercial" selected="true">
        <li class='group'> COMMERCIAL GLASS REPORT</li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
		<li class='group'>Choose Completed Window items that have been SHIPPED 
	<input type='submit' value = 'Remove Selected Items' onClick="Complete.action='glassRemoveConf.asp?ticket=Commercial'; Complete.submit()"></li>
	<li class='group'>Choose Entire POs that have been SHIPPED 
	<input type='submit' value = 'Remove Entire PO' class = "blueButton" onClick="Complete.action='glassRemovePOConf.asp?ticket=Commercial'; Complete.submit()"></li>

<% response.write "<li><table border='1' class='sortable'><tr><th></th><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>SP Colour</th><th>Notes</th><th>Status</th><th>TimeLine</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td><input type='checkbox' name='GID' value='" & RS("ID")& "'></td>"
	response.write "<td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td><td>" & RS("SPColour") & "</td><td>" & RS("NOTES") & "</td>"
%>
<!--#include file="GlassStatus.inc"-->
<%
	Response.write "<td>" & Status & "</td> " ' NStatus
	response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=commercial' target ='#_blank' >Time Line</a> </td>"

	response.write " </tr>"
	rs.movenext
loop
response.write "</table></ul>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

    </ul>
	<input type"hidden" id="ticket" name="ticket" value="Commercial" />
		</form>

</body>
</html>
