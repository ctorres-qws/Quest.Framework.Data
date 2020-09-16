                 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

              <form id="enter" title="Admin Tools" class="panel" name="enter" action="masterin.asp" method="GET" target="_self" selected="true">
              
                              


        <h2>Add to Master List</h2>
  
                       
                                <fieldset>


            <div class="row">
                <label>Q Part #</label>
                <input type="text" name='part' id='part' >
            </div>

			<div class="row">
                <label>Description</label>
                <input type="text" name='description' id='description' >
            </div>
			
			<div class="row">
                <label>Inventory Type</label>
                <select name="inventorytype">
					<option value="Extrusion">Extrusion</option>
					<option value="Gasket">Gasket</option>
					<option value="Hardware">Hardware</option>
					<option value="Plastic">Plastic</option>
					<option value="Sheet">Sheet</option>
					<option value="NPrep FG">NPrep FG</option>
				</select>
            </div>

            <div class="row">
                <label>Supplier #</label>
                <input type="text" name='supplierpart' id='supplierpart' >
            </div>

            <div class="row">
                <label>KG/M</label>
                <input type="text" name='kgm' id='kgm' value = '0'>
            </div>
            
			
			<div class="row">
                <label>Min Stock #</label>
                <input type="number" name='MinLevel' id='MinLevel' value = '250'>
            </div>
			
            <div class="row">
                <label>Can-Art</label>
                <input type="text" name='CanArt' id='CanArt' >
            </div>
			<div class="row">
                <label>HYDRO</label>
                <input type="text" name='HYDRO' id='HYDRO' >
            </div>
			<div class="row">
                <label>KeyMark</label>
                <input type="text" name='Keymark' id='Keymark' >
            </div>
			 <div class="row">
                <label>Extal</label>
                <input type="text" name='Extal' id='Extal' >
            </div>
            
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>
  </form>
                
             
               
</body>
</html>
