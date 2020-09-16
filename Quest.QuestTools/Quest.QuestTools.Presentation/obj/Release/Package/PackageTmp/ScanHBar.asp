         
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
           
			<!-- Matching Program Scan Hbar individual Labels into large Labels-->
			<!-- Designed to ensure every piece of HBar gets scanned before going to Shipping-->
			<!--ScanHbar.asp - ScanHbarMatching.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>H-Bar Match</title>
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
        <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Glazing Scan" class="panel" name="igline" action="ScanHBarMatching.asp" method="GET" selected="true">
         <h2>Bundle Label</h2>

			<div class="row">
                <label>Scan Bundle Label to Continue:</label>
              
            </div>
        <fieldset>
       
             <div class="row">
                <label>Bundle</label>
                <input type="text" name='HBarLabel' id='inputbcw' >
            </div>

     <script type="text/javascript">
		function callback1(barcode) {
			var barcodeText = "BARCODE:" + barcode;
			document.getElementById('barcode').innerHTML = barcodeText;
            console.log(barcodeText); 
        }

		function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
			var textbox = document.getElementById("inputbcw");

			textbox.value = barcode;
    
}
			


        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
         
            </form>
	

</body>
</html>