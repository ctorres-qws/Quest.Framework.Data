<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Hardware Inventory Added April 2017 At request of Lev Bedoev -->
<!--Hardware Master Report Collects from Y_HARDWARE_MASTER-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Hardware Inventory</title>
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
		$('#Inventory').DataTable({
			"iDisplayLength": 25
		});
	});
  
  </script>	
	

<% 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER ORDER BY InventoryType, PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>

<ul id="screen1" title="View All Dies" selected="true">

<%

response.write "<li class='group'>All Master Inventory Information </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='Inventory' id='Inventory'><thead><tr><th>Part</th><th>Description</th><th>HYDRO</th><th>CanArt</th><th>Keymark</th><th>Extal</th><th>KGM</th><th>LBF</th><th>Category</th><th>Minimum Warehoue Level</th><th>InventoryType</th><th>Picture</th></tr></thead><tbody>"
do while not rs.eof
response.write "<tr>"
response.write "<td>" &  rs.fields("PART") & "</td>"
response.write "<td>" &  rs.fields("DESCRIPTION") & "</td>"
response.write "<td>" &  rs.fields("HYDRO") & "</td>"
response.write "<td>" &  rs.fields("Canart") & "</td>"
response.write "<td>" &  rs.fields("Keymark") & "</td>"
response.write "<td>" &  rs.fields("Extal") & "</td>"
response.write "<td>" &  rs.fields("KGM") & "</td>"
response.write "<td>" &  rs.fields("LBF") & "</td>"
response.write "<td>" &  rs.fields("CATEGORY") & "</td>"
response.write "<td>" &  rs.fields("MinLevel") & "</td>"
response.write "<td>" &  rs.fields("InventoryType") & "</td>"
Response.write "<td><img width = '50' height = '50' src='/partpic/" & rs.fields("PART") & ".png'/></td>"
'response.write "<td>" &  rs.fields("Supplierpart") & "</td>"
'response.write "<td>" &  rs.fields("PAINTCAT") & "</td>"

response.write "</tr>"

rs.movenext
loop
RESPONSE.WRITE  "</tbody></table>"
RESPONSE.WRITE "</UL>"


rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
%>

</body>
</html>

