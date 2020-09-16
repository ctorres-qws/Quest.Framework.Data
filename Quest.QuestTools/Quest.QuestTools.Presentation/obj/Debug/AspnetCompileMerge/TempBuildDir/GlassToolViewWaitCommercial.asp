<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Optimization Log Information presented in Report form-->
<!-- Reuqested by Victor and designed by Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Progress</title>
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
<style>
table{
zoom: 70%;
};
 </style>

<!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

	<script type="text/javascript" language="javascript" class="init">

$(document).ready(function() {
	$('#color').DataTable();

} );

	</script>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT top 2000 * FROM Z_GLASSDB WHERE Department = 'Commercial' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>

    </head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

        <form id="Report" title="View Filtered Glass" class="panel" action="GlassToolViewWaitCommercial.asp" name="Report"  method="GET" target="_self" selected="true" > 
	<fieldset>

		<div class="row">
                 <label>Filter </label>
			<select name='FilterType' id='FilterType'>
				<option value="All Active">All Active</option>
				<option value="ALL">All</option>
				<option value="Entered">Entered</option>
				<option value="Optimized">Optimized</option>
				<option value="Ordered">Ordered</option>
				<option value="Completed">Completed</option>
				<option value="Shipped">Shipped</option>

			</select>
		</div>
               <a class="whiteButton" onClick=" Report.submit()">View Filter</a><BR>  

</fieldset>
        <ul id="Profiles" title="Optima View - Report" selected="true">
<% 

FilterType = Request.QueryString("FilterType")
if FilterType = "" then
	FilterType = "ALL Active"
end if

response.write "<li class='group'>Glass Progress -"

Select Case FilterType
	Case "ALL"
	response.write " ALL Glass </li>"
	Case "Entered"

	rs.filter = " OPTIMADATE = NULL AND completedDATE = NULL AND SHIPDATE = NULL"
	response.write " Entered Glass </li>"
	Case "Optimized"
	rs.filter = "OPTIMADATE <> NULL AND EXTRECEIVED = NULL AND INTRECEIVED = NULL AND COMPLETEDDATE=NULL AND SHIPDATE=NULL"
	response.write " Optimized Glass </li>"
	Case "Ordered"
	rs.filter = "OPTIMADATE <> NULL AND COMPLETEDDATE=NULL AND SHIPDATE=NULL"
	Case "Completed"
	rs.filter = "CompletedDATE<>NULL AND ShipDATE = NULL"
	Case "Shipped"
	rs.filter = "SHIPDATE <> NULL"
	Case "All Active"
	rs.filter = "SHIPDATE = NULL"
End Select

response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='color' id ='color' ><thead><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th width ='8pt'>QT File Name</th><th>Order Date</th><th>Status</th><th>Notes</th></thead><tbody>"

Do While Not rs.eof
	Response.write "<tr>"
	Response.write "<td>" & rs.fields("id") & "</td> "
	Response.write "<td>" & rs.fields("JOB") & "</td> " ' Job
	Response.Write "<td>" & rs.fields("FLOOR") & "</td> " ' Floor
	Response.write "<td>" & rs.fields("TAG") & "</td> " ' Tag
	Response.write "<td>" & rs.fields("DIM X") & "</td> " ' Width
	Response.write "<td>" & rs.fields("DIM Y") & "</td> " ' Height
	Response.write "<td>" & rs.fields("1 MAT") & "</td> " ' 1 MATERIAL
	Response.write "<td>" & rs.fields("1 SPAC") & "</td> " ' 1 SPACER
	Response.write "<td>" & rs.fields("2 MAT") & "</td> " ' 2 MATERIAL
	Response.write "<td>" & rs.fields("Department") & "</td> " 'Department
	Response.write "<td>" & rs.fields("ORDERBY") & "</td> " ' Ordered By
	Response.write "<td>" & rs.fields("PO") & "</td> " ' Po Number
	Response.write "<td>" & rs.fields("QTFILE") & "</td> " ' QTFILE
	Response.write "<td>" & rs.fields("InputDate") & "</td> " ' InputDate
%>
<!--#include file="GlassStatus.inc"-->
<%
	Response.write "<td>" & Status & "</td> " ' NStatus
	Response.write "<td>" & rs.fields("NOTES") & "</td> " ' Notes
	Response.write "<td><a href='GlassManageForm.asp?gid=" & rs.fields("ID") & "&ticket=waiting' target='_self' >Manage Glass</a> </td>" 
	Response.write "<td><a href='GlassManageTimeLineForm.asp?gid=" & rs.fields("ID") & "&ticket=waiting' target='_self' >Manage Time Line</a> </td>" 
	Response.write "</tr>"

	rs.movenext
Loop
Response.write "</tbody></table>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

    </ul>
</form>

</body>
</html>
