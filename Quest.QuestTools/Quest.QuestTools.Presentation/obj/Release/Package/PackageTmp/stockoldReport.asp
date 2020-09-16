<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Requested by Shaun Levy and Lev Bedeov - This report displays items in Goreway that were entered more than 3 months ago-->
<!-- This report will help clear up old items in the inventory - Relies on Warehouse Goreway and Input Date -->
<!-- Future Version can also compare Original Value? -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock 3 Months Old</title>
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
    
    <%

		oldDate = DateAdd("m",-3,Date())
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE (DATEIN < #" & oldDate & "# OR DATEIN = NULL) AND Warehouse = 'GOREWAY' ORDER BY PART ASC, DATEIN DESC "
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Goreway Stock Older Than <% response.write oldDate %>" selected="true">
        
<% 

response.write "<li class='group'>Old Stock</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Date Entered</th><th>Qty</th><th>Original Qty</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th></tr>"


do while not rs.eof


Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=intodaytable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & rs("DateIn") & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("firstqty") & " </td>"
Response.write "<td>" & rs("Aisle") & " </td>"
Response.write "<td>" & rs("Rack") & " </td>"
Response.write "<td>" & rs("Shelf") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td><a href='allbypobundle.asp?pobundle=" & rs("PO") & "&ticket=old' target='_self'>" & rs("PO") & "</a></td>"
Response.write "<td><a href='allbypobundle.asp?pobundle=" & rs("Bundle") & "&ticket=old' target='_self'>" & rs("Bundle") & "</a></td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      <li>//END//</li>
	  </ul>                 
          
</body>
</html>
