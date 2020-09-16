
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--Create Labels for SHipping truck using the format  00JOBFLOOR-TAG:DESC-->
		 <!-- Created January 19, 2018 by Michael Bernholtz for Alex Sofienko and Jody Cash-->
		 <!-- Sends information to ShippingOtherLabel_v1-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Sheet Bundle Label</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
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
                <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self" >Panel</a>
    </div>
    
        
    
              <form id="edit" title="Panel Job Name" class="panel" name="edit" action="LabelSheet_v1.asp" method="GET" target="_self" selected="true" >
        <h2>Enter Infomation about Non Window Shipping Item</h2>
  
   
	<br>
	<table align = "center">
		<TR>
			<td align= "Center" width ="200">
			<label >Colour Code</label>
			<input type="text" name='ColorCode' id='ColorCode'>
			</td>
			<td align= "Center" width ="200">
			<label>Job Code</label>
			<input type="text" name='JobCode' id='JobCode'>
			</td>
			<td align= "Center" width ="200">
			<label>ID</label>
			<input type="text" name='SID' id='SID'>
			</td>
		</TR>

		<TR>
			<td align= "Center" width ="200">
			<label >Work Order</label>
			<input type="text" name='WorkOrder' id='WorkOrder'>
			</td>
			<td align= "Center" width ="200">
			<label>Purchase Order</label>
			<input type="text" name='PurchaseOrder' id='PurchaseOrder'>
			</td>
			<td align= "Center" width ="200">
			<label>Date</label>
			<input type="text" name='DateIn' id='DateIn'>
			</td>
		</TR>
		
		<TR>
			<td align= "Center" width ="200">
			<label >Part Number</label>
			<input type="text" name='PartNumber' id='PartNumber'>
			</td>
			<td align= "Center" 
			<label>Size  [ Width by Height ]</label>
			<input type="text" name='size' id='Size' /> 
			</td>			
			<td align= "Center" width ="200">
			<label>Qty</label>
			<input type="text" name='Qty' id='Qty'>
			</td>
		</TR>
	</table>
		
	</div>    

        <BR>
		<a class="lightblueButton" onClick="edit.submit()">Create Label</a><BR>
        
       </form>
         
    
</body>
</html>



