<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--Create Labels for SHipping truck using the format  00JOBFLOOR-TAG:DESC-->
		 <!-- Created January 19, 2018 by Michael Bernholtz for Alex Sofienko and Jody Cash-->
		 <!-- Sends information to ShippingOtherLabel_v1-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping Label (Non-Window)</title>
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
    
        
    
              <form id="edit" title="Panel Job Name" class="panel" name="edit" action="ShippingOtherLabel_v1.asp" method="GET" target="_self" selected="true" >
        <h2>Enter Infomation about Non Window Shipping Item</h2>
  
   

<fieldset>
	<div class="row">
		<label>Job</label>
		<input type="text" name='Job_Name' id='Job_Name'>
	</div>
	<div class="row">
		<label>Floor</label>
		<input type="text" name='Floor_Name' id='Floor_Name'>
	</div>
	<div class="row">
		<label>Tag</label>
		<input type="text" name='Tag_Name' id='Tag_Name'>
	</div>
	<div class="row">
		<label>QTY</label>
		<input type="text" name='Qty_Name' id='Qty_Name'>
	</div>
	<div class="row">
		<label>Description</label>
		<input type="text" name='Desc_Name' id='Desc_Name'>
	</div>
            
            
</fieldset>


        <BR>
		<a class="lightblueButton" onClick="edit.submit()">Create Label</a><BR>
        
       </form>
<% 
DBConnection.close
set DBConnection=nothing
%>           
    
</body>
</html>



