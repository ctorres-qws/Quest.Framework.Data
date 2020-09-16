<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Scan Page to Reduce Sheet Inventory-->

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
        <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Panel</a>
    </div>

    <form id="enter" title="Use Inventory" class="panel" name="enter" action="sheetreduceconf.asp" method="GET" target="_self" selected="true">
              
	
	<h2>Remove Sheets from Inventory</h2>
		<fieldset>	
		<div class="row">
			<label>Job</label>
			<input type="text" name='Job' id='Job' >
		</div>
		<div class="row">
			<label>Thickness</label>
			<select name="Thickness">
				<option value= "0.27" >0.27 </option>
				<option value= "0.50" >0.50 </option>
				<option value= "0.63" >0.63 </option>
				<option value= "0.80" >0.80</option>
				<option value= "0.125" >0.125</option>
			</select>
        </div>
		<div class="row">
			<label>Qty</label>
			<input type="text" name='Qty' id='Qty' >
		</div>
	</fieldset>
	<BR>
	<a class="whiteButton" href="javascript:enter.submit()">Submit</a>

</form>    
<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("Job");
        textbox.value = barcode;

}
    
        </script>
      
</body>
</html>
