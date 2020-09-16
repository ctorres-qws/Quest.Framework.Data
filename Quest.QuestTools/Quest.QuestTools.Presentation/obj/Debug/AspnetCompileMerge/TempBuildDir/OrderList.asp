                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Order Glass</title>
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
    $('#GlassOrder').DataTable();
} );
  
  </script>
 <style>
table{
zoom: 80%
};
 </style>

    <%
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM GlassOrder Where [Active] = True ORDER BY PO ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Order" target="_self">Order Entry</a>
        </div>
      
       
        <ul id="Profiles" title=" Glass Orders" selected="true">
        <li class='group'>Glass Orders </li>
         <a class="whiteButton" href="OrderListEnter.asp" target='_Self'>Add New Order</a>
<% 

response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='GlassOrder' id='GlassOrder'><thead><tr><th width = '120'>PO</th><th>Glass Code</th><th>Job</th><th>Floor</th><th>Quantity</th><th>From</th><th>Order By</th><th>Ship to QT</th><th>Order Date</th><th>Expected Date</th><th width ='40'>Notes</th><th>Acknowledged</th><th>Broken</th></th><th width ='150'>Manage</th></tr></thead><tbody>"
if rs.eof then
Response.write "<tr><td colspan ='14'>No current orders</td></tr>"
end if	
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("PO") & "</td>"
	response.write "<td>" & RS("GlassCode") &"</td>"
	response.write "<td>" & RS("JOB") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("Qty") & "</td>"
	response.write "<td>" & RS("From") & "</td>"
	response.write "<td>" & RS("OrderBy") & "</td>"
	response.write "<td>" & RS("ShipOutDate") & "</td>"
	response.write "<td>" & RS("OrderDate") & "</td>"
	response.write "<td>" & RS("ExpectedDate") & "</td>"
	response.write "<td>" & RS("Notes") & "</td>"
	response.write "<td>" & RS("Ack") & "</td>"
	response.write "<td>" & RS("Broken") & "</td>"
	response.write "<td> <a class='lightblueButton' href='OrderListEditForm.asp?ORID=" & RS("OrderID") & "' target='_Self'>Edit this Order</a></td>"
	response.write " </tr>"

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
