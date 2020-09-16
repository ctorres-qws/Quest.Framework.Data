<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 11th, by Michael Bernholtz - View Page for all items that are marked Broken-->
<!-- Form created at Request of Ariel Aziza Implemented by Michael Bernholtz--> 
<!-- Using Tables: X_Broken -->

<!-- Inputs to ShippingItemPrintLabel.asp.asp-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Broken Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script src="sorttable.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    

<%
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_Broken WHERE orderDate is NULL ORDER BY Job, Floor ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection



%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Broken Glass </h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
    </div>
	
	<ul id="BrokenGlass" title=' Broken Glass' selected='true'>
	<li class ="group"> <a href="BrokenGlassReportAll.asp" target = "_self">Show All (Inclduing Ordered)</a></li>
	<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Opening</th><th>Width</th><th>Height</th><th>Added By</th><th>Added Date</th><th>Order #</th><th>Ordered Date</th><th>Reason</th><th>Notes</th></tr>
		
<%

		do while not rs.eof
				Response.write "<tr>"
				Response.write "<td>" & trim(RS("job")) & "</td><td>" & trim(RS("floor")) & "</td><td>" & trim(RS("tag")) & "</td>"
				Response.write "<td>" & trim(RS("opening")) & "</td><td>" & trim(RS("width")) & "</td><td>" & trim(RS("height")) & "</td>"
				Response.write "<td>" & trim(RS("addby")) & "</td><td>" & trim(RS("addDate")) & "</td><td>" & trim(RS("ordernum")) & "</td>"
				Response.write "<td>" & trim(RS("orderDate")) & "</td><td>" & trim(RS("Reason")) & "</td><td>" & trim(RS("Notes")) & "</td>"
				Response.write "</tr>"
		rs.movenext
		loop
		response.write "</table></li>"
		rs.close
		set rs = nothing
		DBConnection.close
		set DBConnection = nothing		


%>
  </ul>          
         
               
</body>
</html>
