<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Full listing of Sapa items - At Request of Ruslan Bedeov, Built by Michael Bernholtz, May 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY WAREHOUSE, PART"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"

response.write "<li class='group'>HYDRO</li>"

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

        <ul id='SAPA' title=' HYDRO Inventory' selected='true'> 
		<li class='group'>HYDRO INVENTORY </li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
		<li><table border='1' class='sortable' width='95%'>
		<tr><th>Part</th><th>Color/Project</th><th>Length (Feet)</th><th>Quantity (SL)</th><th>PO</th><th>Bundle</th><th>Length (mm)</th><th>Manage</th></tr>
        
<% 

do while not rs.eof
Response.write "<tr><td>" & rs.fields("PART") & "</td>"
Response.write "<td>" & rs.fields("Colour") & "</td>"

Response.write "<td> " & rs.fields("Lft") & "'</td>"
Response.write "<td> " & rs.fields("Qty") & " </td>"
Response.write "<td> " &  rs.fields("po") & " </td>"
Response.write "<td> " &  rs.fields("bundle") & " </td>"
Response.write "<td> " & rs.fields("Lmm") & "mm </td>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=stockhydro&id=" & rs("id") & "' target='_self'>Manage</a></td>"

Response.write "</tr>"

rs.movenext
loop

RESPONSE.WRITE "</table></li>"

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
