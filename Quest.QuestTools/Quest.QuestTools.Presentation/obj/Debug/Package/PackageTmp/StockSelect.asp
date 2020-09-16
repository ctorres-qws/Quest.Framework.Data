<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Created July 16th, by Michael Bernholtz - Select Choice for Stock Entry-->
<!-- Stock in Y_INV now includes many items, not just Extrusions at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Sends to stock.asp#_enter-->
<!--#include file="countrylocation.inc"-->
<%
ScanMode = True
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Type</title>
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
        <h1 id="pageTitle">Choose Stock Type</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>

 <form id="stockType" title="Select Inventory Type" class="panel" name="stockType" method="GET" target="_self" selected="true">
              
			  <h2>Choose Inventory Type</h2>
              <fieldset>

			<div class="row">
			<label>Choose the Inventory Type from the Dropdown</label>
			<select id='InventoryType' name="InventoryType" >
				<option value="Extrusion">Extrusion</option>
				<option value="Sheet">Sheet</option>			
				<option value="Gasket">Gasket</option>
				<option value="Hardware">Hardware</option>
				<option value="Plastic">Plastic</option>
			<!--	<option value="NPrep FG">NPrep FG</option> -->
			</select>

            </div>

        		<a class="whiteButton" onClick="stockType.action='stock.asp#_enter'; stockType.submit()">Add Inventory Items</a><BR>
			</fieldset>
</form>

</body>
</html>
