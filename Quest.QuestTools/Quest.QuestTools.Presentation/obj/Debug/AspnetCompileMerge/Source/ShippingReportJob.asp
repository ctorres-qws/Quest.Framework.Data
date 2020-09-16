<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko -->
<!-- Finds the Job and Floor Requirements and matches it to the SHipping -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping View</title>
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
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self" >Report</a>
    </div>
    
        
    
              <form id="edit" title="Select Job and Floor" class="panel" name="edit" action="ShippingReportJob1.asp" method="GET" target="_self" selected="true" > 
        <h2>Select a Job and a Floor</h2>
  

<fieldset>
     <div class="row">
                <label>JOB</label>
                <input type="text" name='job' id='job'>
    </div>
    <div class="row">
                <label>Floor</label>
                <input type="text" name='floor' id='floor'>
    </div>       
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Show all Items and Shipping Status</a><BR>

            </form> 
</body>
</html>
