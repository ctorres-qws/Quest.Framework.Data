<!--#include file="@common.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!-- New Glazing Scan Coded Feb 2016 to allow specific openings to be marked completed -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>CDN Glazing</title>
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
   
 <%
 Employee = request.querystring("EmployeeID")
 %>
   
   
    <form id="igline" title="Glazing Scan" class="panel" name="igline" action="ScanGlazingWindow.asp" method="GET" selected="true">
         <h2>NEW Glazing Scan</h2>

			<div class="row">
                <label>Scan Window and Employee to continue:</label>
              
            </div>
        <fieldset>
       
             <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
			
			
			<div class="row">
				<label>Employee</label>
				<select name= 'empid' id = 'empid'>
					<option value="1111" <% if Employee = 1111 then response.write "Selected" %>>Line 1 Day</option>
					<option value="1000" <% if Employee = 1000 then response.write "Selected" %>>Line 1 Night</option>
					<option value="2222" <% if Employee = 2222 then response.write "Selected" %>>Line 2 Day</option>
					<option value="2000" <% if Employee = 2000 then response.write "Selected" %>>Line 2 Night</option>
					<option value="3333" <% if Employee = 3333 then response.write "Selected" %>>Line 3 Day</option>
					<option value="3000" <% if Employee = 3000 then response.write "Selected" %>>Line 3 Night</option>
					<option value="9999" <% if Employee = 9999 then response.write "Selected" %>>Panel Day</option>
					<option value="9000" <% if Employee = 9000 then response.write "Selected" %>>Panel Night</option>
					<option value="8888" <% if Employee = 8888 then response.write "Selected" %>>R3/Shift Day</option>
					<option value="8000" <% if Employee = 8000 then response.write "Selected" %>>R3/Shift Night</option>
				</select>
			</div>	

     <script type="text/javascript">
		function callback1(barcode) {
			var barcodeText = "BARCODE:" + barcode;
			document.getElementById('barcode').innerHTML = barcodeText;
            console.log(barcodeText); 
        }

		function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
			var textbox = document.getElementById("inputbcw");
			if ( barcode.length == 4 ) {
				textbox = document.getElementById("empid");
			}
    
			textbox.value = barcode;
    
}
			


        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
         
            </form>
	

</body>
</html>