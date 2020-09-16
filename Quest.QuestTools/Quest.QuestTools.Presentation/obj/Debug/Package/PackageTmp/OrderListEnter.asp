                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--#include file="dbpath.asp"--> 
<!--Add new item to orderlist-->
<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>New Order</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="OrderList.asp" target="_self">Order List</a>
        </div>

            
              <form id="enter" title="Enter New Order" class="panel" name="enter" action="OrderListconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Order:</h2>
  
                       
       <fieldset>

	<div class="row">
		<label>PO</label>
		<input type="text" name='PO' id='PO' >
	</div>

	<div class="row">
		<label>Glass Code</label>
		<select name="GlassCode">
			<% mat = mat1 %>
			<!--#include file="QSU.inc"-->
		</select>
    </div>

    <div class="row">
		<label>Job </label>
		<select name="Job">
		<option value = "" selected>-</option>
			<% ActiveOnly = True %>
			<!--#include file="JobsList.inc"-->
		</select>
    </div>
	<div class="row">
		<label>Floor </label>
		<input type="text" name='Floor' id='Floor' >
    </div>
	<div class="row">
		<label>Quantity</label>
		<input type="number" name='Qty' id='Qty' value = "0">
    </div>
	<div class="row">
		<label>Glass Ordered From</label>
        <select name= 'From' id = 'From'>
			<option value="QuickTemp">QuickTemp</option>
			<option value="Cardinal">Cardinal</option>
			<option value="Woodbridge">Woodbridge</option>
			<option value="Saand">Saand</option>
			<option value="TruLite">TruLite</option>
		</select>
     </div>	
	
	<div class="row">
		<label>Ordered By</label>
		<input type="text" name='OrderBy' id='orderBy' >
    </div>
	
	<div class="row">
		<label>Ship to QT</label>
		<input type="date" name='ShipOutDate' id='ShipOutDate' >
    </div>
	<div class="row">
		<label>Order Date</label>
		<input type="date" name='orderDate' id='orderDate' >
    </div>
	<div class="row">
		<label>Expected Date</label>
		<input type="date" name='ExpectedDate' id='ExpectedDate' >
    </div>
	
	<div class="row">
		<label>Notes </label>
		<input type="text" name='Notes' id='Notes' >
    </div>

	
            
                    <a class="whiteButton" href="javascript:enter.submit()" target='_Self'>Submit</a>
            
         
</fieldset>


            
            </form>
                
  <%

rs5.close
set rs5 = nothing
rsJob.close
set rsJob = nothing
DBConnection.close 
set DBConnection = nothing
%>  
               
</body>
</html>
