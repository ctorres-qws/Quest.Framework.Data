<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 5th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Mark Done</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
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
	$('#Done').dataTable({
		paging: false
	});
});

</script>
<!-- Css code to shrink the table to stay on one page with extra moz code for firefox -->
<style>

table {
	zoom: 80%;
	-moz-transform: scale(0.85);
	-webkit-transform: scale(1.0);
	-moz-transform-origin: left top;
};

</style>
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

        <ul id="Profiles" title="Mark Glass Order Done" selected="true">
<%

	MARKEDID = REQUEST.QueryString("ID")
	If MARKEDID = "" Then
		MARKEDID = -1
	End If
	UNDO = REQUEST.QueryString("UNDO")
	OUTDATE = Date

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Call Process(false, true)
	Case c_MODE_HYBRID
		Call Process(false, true)
		Call Process(true, false)
	Case c_MODE_SQL_SERVER
		Call Process(true, true)
End Select

Function Process(isSQLServer, b_Display)
	DBOpen DBConnection, isSQLServer

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID DESC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	'Undo button only appears after an item is marked done
	Select Case UNDO
		Case 0
			rs.Filter = "ID = " & CLng(MARKEDID)
			If not rs.eof Then
				rs.Fields("COMPLETEDDATE") = NULL
				rs.Fields("Hide") = NULL
				rs.update
			End If
			rs.Filter = ""
		Case 1
			If b_Display Then Response.Write "<li class='group'><a href='glassmarkdonetable.asp?ID=" & MARKEDID & "&UNDO=0' target='_self'> UNDO - " & MARKEDID & "</a></li>"
			rs.Filter = "ID = " & CLng(MARKEDID)
			If not rs.eof Then
				rs.Fields("COMPLETEDDATE") = OUTDATE
				rs.Fields("Hide") = "Completed"
				rs.update
			End If
			rs.Filter = ""
	End Select

	DbCloseAll

End Function

Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp
DBConnection.Open DSN

Set rs = Server.CreateObject("adodb.recordset")
Rs.open "select * FROM Z_GLASSDB AND ISNULL(COMPLETEDDATE) ORDER BY ID ASC",DBConnection,1,3

Response.write "<li>Select a Glass Item to Mark Done:</li>"

%>
<li><table border='1' class="Done" id="Done" width = "100%"> 
<thead><tr><th>ID</th><th width ="10">Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th>
<th>1 Mat</th><th>2 Mat</th><th>Input Date</th><th>Required Date</th>
<th>Type</th><th>Order</th><th>PO</th><th width ="30">Note</th><th width ="50">Mark Done</th></tr></thead><tbody>
<%
Do While Not rs.eof
		Response.write "<tr>"
		Response.write"<td>" & rs("ID") & "</td>"
		Response.write"<td>" & rs("JOB") & "</td>"
		Response.write"<td>" & rs("FLOOR") & "</td>"
		Response.write"<td>" & rs("TAG") & "</td>"
		Response.write"<td>" & rs("Dim X") & "</td>"
		Response.write"<td>" & rs("Dim Y") & "</td>"
		Response.write"<td>" & rs("1 MAT") & "</td>"
		Response.write"<td>" & rs("2 MAT") & "</td>"
		Response.write"<td>" & rs("INPUTDATE") & "</td>"
		Response.write"<td>" & rs("REQUIREDDATE") & "</td>"
		Response.write"<td>" & rs("DEPARTMENT") & "</td>"
		Response.write"<td>" & rs("ORDERBY") & "</td>"
		Response.write"<td>" & rs("PO") & "</td>"
		Response.write"<td>" & rs("NOTES") & "</td>"
		Response.write"<td><a class='greenButton' href='glassmarkdonetable.asp?ID=" & rs("ID") & "&UNDO=1' target='_self'>Completed</a> </td>"
		Response.write "</tr>"
	rs.movenext
Loop
Response.write "</tbody></table></li>"
%>
      </ul>
<%
rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
%>
</body>
</html>
