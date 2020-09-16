<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Optimization Log Information presented in Report form-->
<!-- Reuqested by Victor and designed by Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optimization Log Report</title>
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
<!-- Fixed Headers -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/extensions/FixedHeader/js/dataTables.fixedHeader.js"></script>
 
  <script type="text/javascript">
  $(document).ready( function () {
    $('#color').DataTable();
} );
  
  </script>

 
 
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM OptimizeLog ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>
 
    </head>
<body>


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_GlassP" target="_self">Glass Prod</a>
        </div>

       
        <ul id="Profiles" title="Optimization Log - Report" selected="true">
<% 
response.write "<li class='group'>Optimization Log</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='color' id ='color' ><thead><tr><th>Job</th><th>Floor</th><th>Glass</th><th>Type</th><th>Opt File</th><th># of Lites</th><th>Bending File</th><th>Opt Date</th><th>Opt Time</th><th>Shift</th><th>Employee</th><th>Date Glass Cut</th><th>Time Glass Cut</th><th>Ship (QuickTemp)</th><th>Received (QuickTemp)</th><th>Back-Order Lites</th><th>Back-Order Names</th><th></th></tr></thead><tbody>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("Glass") & "</td>"
	response.write "<td>" & RS("Type") & "</td>"
	response.write "<td>" & RS("OpFile") & "</td>"
	response.write "<td>" & RS("Lites") & "</td>"
	response.write "<td>" & RS("BendFile") & "</td>"
	response.write "<td>" & RS("OpDate") & "</td>"
	response.write "<td>" & RS("OpTime") & "</td>"
	response.write "<td>" & RS("Shift") & "</td>"
	response.write "<td>" & RS("Employee") & "</td>"
	response.write "<td>" & RS("GlassCutDate") & "</td>"
	response.write "<td>" & RS("GlassCutTime") & "</td>"
	response.write "<td>" & RS("ShipDate") & "</td>"
	response.write "<td>" & RS("ReceivedDate") & "</td>"
	response.write "<td>" & RS("BackOrder") & "</td>"
	response.write "<td>" & RS("BackOrderText") & "</td>"

	Response.write "<td><a href='OptimizationLogEditForm.asp?OPID=" & rs.fields("ID") & "' target='_self' >Manage</a> </td>" 
	response.write "</tr>"
	rs.movenext
loop
response.write "</tbody></table>"


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>        
            
       
               
</body>
</html>
