<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Created January 2018, by Michael Bernholtz - View All items on All Trucks with the same name-->
<!-- Entry page for ShippingTruckView.asp report -->
<!-- Sokol Requested the ability to see all items on multiple trucks that all have the same name-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>View Items on Truck</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script src="sorttable.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
   

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">Truck View</h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Report" target="_self">Report</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="AddTruck" title="Add Truck" class="panel" name="AddTruck" action="ShippingTruckView.asp" method="GET" selected="true">
        
        <h2>Input Truck Name</h2>
       <fieldset>
	  
		<div class="row">
                <label>Truck</label>
                <input type="text" name='Truck' id='Truck' >
        </div>
		</fieldset>
	
		<a class="whiteButton" onClick="AddTruck.submit()">Submit</a><BR>
            </form>
		

</body>
</html>