<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created July 2019 - by Michael Bernholtz --> 
<!--Delete Scan Page for Scan Items-->
<!-- No manual Entry, Only Scan Reuqirement by David Ofir-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Delete Ship Scan</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  				<script type="text/javascript">
				function FlashingScreen() {

						var flash = false;
						var task = setInterval(function() {
						if(flash = !flash) {
							document.body.style.backgroundColor = '#0ff';
							document.panel.style.backgroundColor = '#0f0';
						} else {
							document.body.style.backgroundColor = '#f00';
						}
						}, 1000);
						
					}
					FlashingScreen()
					
			  </script>
  

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">Unscan Confirmation</h1>
        <a class="button leftButton" type="cancel" href="ShipHomeManager.HTML" target="_self">Scan</a>
     </div>
   
   
   
    <form id="igline" title="UNSCAN SCAN" class="panel3" name="igline" action="ShipTruckScanDeleteConf.asp" method="GET" selected="true">
         <h2>Remove Window</h2>
		 	
        <fieldset>
       
	   
	        <div class="row">
                <label>Scan </label>
                <input type="text" name='DeleteID' id='inputbcw' value = '' readonly>
            </div>
	   
            
            </fieldset>
			 <BR>
            </form>
			
			
			

<script type="text/javascript">
				  
				 		  
	function callback1(barcode) {
		var barcodeText = "BARCODE:" + barcode;

		document.getElementById('barcode').innerHTML = barcodeText;
		console.log(barcodeText);
        
	}
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
   
        textbox.value = barcode;
		igline.submit();

	}
    
</script>


</body>
</html>