<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Optimization Log Information presented in Report form-->
<!-- Reuqested by Victor and designed by Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Ordered Items</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

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
// Tabs
	$('#tabs').tabs();

} );

				
	</script>

 
 

 
    </head>
<body>
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE OPTIMADate = NULL ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

       
        <ul id="Profiles" title="Optimization Log - Report" selected="true">
		<div id="tabs">
			<ul>
				<li><a href="#tabs-1">Details</a></li>
				<li><a href="#tabs-2">Timeline</a></li>
				<li><a href="#tabs-3">Additional Info</a></li>
			</ul>
				
		
		
<% 
response.write "<div id='tabs-1'>"
response.write "<li class='group'>Optimization Log</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='color' id ='color' ><thead><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Orderby</th><th>PO</th><th>1 Mat</th><th>1 Spac</th><th>2 Mat/th><th>Notes</th><th>Details</th><th>Timeline</th></tr></thead><tbody>"
do while not rs.eof

	response.write "<tr>"
		Response.write "<td> " & rs.fields("id") & "</td> "
		Response.write "<td>" & rs.fields("JOB") & "</td> " ' Job
		Response.Write "<td>" & rs.fields("FLOOR") & "</td> " ' Floor
		Response.write "<td>" & rs.fields("TAG") & "</td> " ' Tag
		Response.write "<td>" & rs.fields("ORDERBY") & "</td> " ' Ordered By
		Response.write "<td>" & rs.fields("PO") & "</td> " ' Po Number
		Response.write "<td>" & rs.fields("1 MAT") & "</td> " ' 1 MATERIAL
		Response.write "<td>" & rs.fields("1 SPAC") & "</td> " ' 1 SPACER
		Response.write "<td>" & rs.fields("2 MAT") & "</td> " ' 2 MATERIAL
		Response.write "<td>" & rs.fields("NOTES") & "</td> " ' Notes
		Response.write "<td><a href='GlassManageForm.asp?gid=" & rs.fields("ID") & "&ticket=active' target='_self' >Manage Glass</a> </td>" 
	Response.write "</tr>"

rs.movenext
loop
response.write "</tbody></table></div>"

response.write "<div id='tabs-2'>"
response.write "<li class='group'>Optimization Log Timeline</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='color' id ='color' ><thead><tr><th>Order Date</th><th>Optima</th><th>Exterior Received</th><th>Interior Received</th><th>Sealed Date</th><th>Ship Date</th><th>Timeline</th></tr></thead><tbody>"
rs.movefirst
do while not rs.eof

	response.write "<tr>"
		Response.write "<td> " & rs.fields("InputDate") & "</td> "
		Response.write "<td>" & rs.fields("OptimaDate") & "</td> " ' Job
		Response.Write "<td>" & rs.fields("ExtReceived") & "</td> " ' Floor
		Response.write "<td>" & rs.fields("IntReceived") & "</td> " ' Tag
		Response.write "<td>" & rs.fields("CompletedDate") & "</td> " ' Ordered By
		Response.write "<td>" & rs.fields("ShipDate") & "</td> " ' Po Number
		Response.write "<td><a href='GlassManageTimeLineForm.asp?gid=" & rs.fields("ID") & "&ticket=active' target='_self' >Manage Time Line</a> </td>" 
	Response.write "</tr>"

rs.movenext
loop
response.write "</tbody></table></div>"
response.write "<div id='tabs-3'>"
response.write "<li class='group'>Optimization Log Timeline</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='color' id ='color' ><thead><tr><th>Order Date</th><th>Optima</th><th>Exterior Received</th><th>Interior Received</th><th>Sealed Date</th><th>Ship Date</th><th>Timeline</th></tr></thead><tbody>"
rs.movefirst
do while not rs.eof

	response.write "<tr>"
		Response.write "<td> " & rs.fields("InputDate") & "</td> "
		Response.write "<td>" & rs.fields("OptimaDate") & "</td> " ' Job
		Response.Write "<td>" & rs.fields("ExtReceived") & "</td> " ' Floor
		Response.write "<td>" & rs.fields("IntReceived") & "</td> " ' Tag
		Response.write "<td>" & rs.fields("CompletedDate") & "</td> " ' Ordered By
		Response.write "<td>" & rs.fields("ShipDate") & "</td> " ' Po Number
		Response.write "<td><a href='GlassManageTimeLineForm.asp?gid=" & rs.fields("ID") & "&ticket=active' target='_self' >Manage Time Line</a> </td>" 
	Response.write "</tr>"

rs.movenext
loop
response.write "</tbody></table></div>"
response.write "</div>"
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>        
            
       
               
</body>
</html>
