<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- TMP STOCK FOR DATA ENTRY FROM HORNER / DURAPAINT-->

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
        <a class="button leftButton" type="cancel" href="index.html#_TmpINV" target="_self">TMP Stock</a>
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
			</select>		
	
            </div>
			
        		<a class="whiteButton" onClick="stockType.action='stockTEMP.asp#_enter'; stockType.submit()">Add Inventory Items</a><BR>
            
         
			</fieldset>
</form>       
               
</body>
</html>
