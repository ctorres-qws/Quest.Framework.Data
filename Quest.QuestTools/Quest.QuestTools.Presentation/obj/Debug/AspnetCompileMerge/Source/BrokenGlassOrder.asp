<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 11th, by Michael Bernholtz - Order Page for all items that are marked Broken-->
<!-- Form created at Request of Ariel Aziza Implemented by Michael Bernholtz--> 
<!-- Using Tables: X_Broken -->
<!--Inputs to BrokenGlassOrderConf.asp-->

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

 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Broken Glass </h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
    </div>
	<form id="order" title="Order Broken Glass" class="panel" name="order" action="BrokenGlassOrder.asp" method="Post" target="_self" selected="true">
           <h2>Type in the Order Number and then select Broken Items purchased on that order</h2>
		   <fieldset>   
		   
		   	<div class="row" >
				<label>Order #</label>
				<input type="text" name='ordernum' id='ordernum' >
			</div>
		   
		   
		<div>
	<table border='1' class='sortable'><tr><th>Ordered</th><th>Job</th><th>Floor</th><th>Tag</th><th>Opening</th><th>Notes</th></tr>
		
<%
		count = 0
		do while not rs.eof
				Response.write "<tr><td><input type='checkbox' name='glass' & value ='" & rs("id") & "'></input></td>"
				Response.write "<td>" & trim(RS("job")) & "</td><td>" & trim(RS("floor")) & "</td><td>" & trim(RS("tag")) & "</td><td>" & trim(RS("opening")) & "</td><td>" & trim(RS("notes")) & "</td></tr>"
				
		rs.movenext
		loop
		response.write "</table>"
		rs.close
		set rs = nothing
		DBConnection.close
		set DBConnection = nothing				


%>

			</fieldset>
			
			<a class="whiteButton" onClick="order.action='BrokenGlassOrderConf.asp'; order.submit()">Submit</a><BR>
			
  </div>         
  </form>       
               
</body>
</html>
