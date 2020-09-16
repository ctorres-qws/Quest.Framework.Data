<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 6th, by Michael Bernholtz - Entry Form to input items to QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->

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
        <h1 id="pageTitle">Add Items to QC Inventory</h1>
        <a class="button leftButton" type="cancel" href="index.html#_QC" target="_self">QC</a>
        </div>
   
   
   
   
            
              <form id="enter" title="Add QC Master Item" class="panel" name="enter" action="QCMasterConf.asp" method="GET" target="_self" selected="true">
              
			  <h2>Add QC Inventory Master Item</h2>
              <fieldset>               


        
		<div class="row">

			<label>Inventory Type</label>
			<select id='InventoryType' name="InventoryType" >
				<option value="QCGlass">Glass</option>
				<option value="QCSpacer">Spacer</option>
				<option value="QCSealant">Sealant</option>
				<option value="QCMisc">Miscellaneous</option>
			</select>
            </div>
                       
        <div class="row">
            <label>Item Name</label>
            <input type="text" name='ItemName' id='ItemName' >
        </div>

        <div class="row">
            <label>Manufacturer</label>
            <input type="text" name='MANUFACTURER' id='MANUFACTURER' >
        </div>
		
		
		<h3> &nbsp &nbsp &nbsp Additional Fields for Glass Only</h3>
	
		<div class="row">
            <label>Code</label>
            <input type="text" name='Code' id='Code' >
        </div>
	<div class="row">
        <label>Pack Count</label>
        <input type="Number" name='Pieces' id='Pieces' value="16" >
    </div>

    <div class="row">
        <label>Width</label>
        <input type="Number" name='Width' id='Width' value="96" >
    </div>
               
    <div class="row">
        <label>Height</label>
        <input type="Height" name='Height' id='Height' value="130" >
    </div> 
	               
    <div class="row">
        <label>$ per Sqft</label>
        <input type="Number" name='Price' id='Price' value="1" >
    </div> 
            
        <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>


            
            </form>
        
             
               
</body>
</html>
