<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit and Delete Form for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- February 2019 - USA tables added-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
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
        <h1 id="pageTitle">Select Glass to Edit</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_QC" target="_self">Glass<%response.write HomeSiteSuffix%></a>
    </div>  


 <form id="QCitem" title="Select QC Inventory Type" class="panel" name="QCitem" method="GET" target="_self" selected="true">

			  <h2>Choose Inventory Type</h2>
              <fieldset>

			<div class="row">
			<label>Choose the Inventory Type from the Dropdown</label>
			<select id='InventoryType' name="InventoryType" >
				<option value="QCGlass">Glass</option>
				<option value="QCSpacer">Spacer</option>
				<option value="QCSealant">Sealant</option>
				<option value="QCMisc">Misc</option>
			</select>
            </div>

        		<a class="whiteButton" onClick="QCitem.action='qcInventoryEdit.asp'; QCitem.submit()">Manage Inventory Items</a><BR>

			</fieldset>
</form>

<%
DBConnection.close
Set DBConnection = nothing
%>

</body>
</html>
