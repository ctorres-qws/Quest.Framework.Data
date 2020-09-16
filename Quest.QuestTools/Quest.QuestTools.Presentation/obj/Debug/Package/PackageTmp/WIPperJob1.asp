<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Inventory and WIP Per Job Floor Requested by Shaun Levy, March 2017 -->
<!-- Job Floor Breakdown leads to 2 pages   Inventory per Job and WIP per Job  -->
<!-- InventoryperJob.asp, WIPperJob.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self" >Inventory</a>
    </div>
    
        
    
              <form id="CDate" title="Date for Stock/Prod" class="panel" name="CDate"  method="GET" target="_self" selected="true" > 
			  <h2>Choose a JOB and FLOOR for Stock Entry or Sent to Production</h2>

<fieldset>
	<div class="row">
		<label>Job</label>
		<input type="text" name='Job' id='Job' /> 
	</div>
	<div class="row">
		<label>Floor</label>
		<input type="text" name='Floor' id='Floor' /> 
	</div>
            
</fieldset>


        <BR>

        <a class="greenButton" onClick="CDate.action='InventoryPerJob.asp'; CDate.submit()">Current Stock</a><BR>
		<a class="greenButton" onClick="CDate.action='WIPperJob.asp'; CDate.submit()">Stock in Production</a><BR>
          
            </form> 
            
</html>


