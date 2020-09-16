<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Updated January 2015, Michael Bernholtz, to split Job and Side rather than a single field. This will help with Database Consistency -->
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
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>
   
   
   
   
            
              <form id="enter"  "title="Add Color" class="panel" name="enter" action="colorin.asp" method="GET" target="_self" selected="true">
              
                              


        <h2>Add A New Color</h2>
		<h2> Please note the Change: Please add the Job to the first Field and select Ext/Int in the dropbox </h2>
		
  
                       
    <fieldset>

        <div class="row">
			<label>JOB Code</label> 
            <input type="text" name='JOB' id='JOB' >
        </div>
		<div class="row">
            <label>Ext / Int</label>
            <Select name='Side'>
				<option value="Ext.">Exterior</option>
				<option value="Int.">Interior</option>
			</Select>
        </div>

		<div class="row">
            <label>Paint Code</label>
            <input type="text" name='CODE' id='CODE' >
        </div>
		
		<div class="row">
            <label>Paint Type</label>
            <select name='Company'>
				<option value="PPG Acrynar">PPG Acrynar</option>
				<option value="PPG Duranar">PPG Duranar</option>
				<option value="PPG Duracron">PPG Duracron</option>
				<option value="PPG Duracron White">PPG Duracron White K-1285</option>
				<option value="PPG Duranar XL">PPG Duranar XL</option>
				<option value="PPG Duranar XL + Basecoat">PPG Duranar XL + Basecoat</option>
				
				<option value="VALSPAR Acrodize">VALSPAR Acrodize</option>
				<option value="VALSPAR Acroflur">VALSPAR Acroflur</option>
				<option value="VALSPAR Clear Anodize">VALSPAR Clear Anodize</option>
				<option value="VALSPAR Fluropon">VALSPAR Fluropon</option>
				<option value="VALSPAR Fluropon Classic">VALSPAR Fluropon Classic</option>
				<option value="VALSPAR Flurospar">VALSPAR Flurospar</option>
				<option value="VALSPAR Polylure">VALSPAR Polylure</option>

				<option value="Other">Other</option>
			</Select>
        </div>

        <div class="row">
            <label>Description</label>
            <input type="text" name='DESCRIPTION' id='DESCRIPTION' >
        </div>
           
        <div class="row">
            <label>Price Cat.</label>
            <input type="text" name='PAINTCAT' id='PAINTCAT' >
        </div>
       
	    <div class="row">
            <label>Active</label>
            <input type="checkbox" name='Active' id='Active' checked>
        </div> 
        
	    <div class="row">
            <label>Extrusion Colour</label>
            <input type="checkbox" name='EXTRUSION' id='EXTRUSION' >
        </div> 

	    <div class="row">
            <label>Sheet Colour</label>
            <input type="checkbox" name='SHEET' id='SHEET' >
        </div> 		
            
            
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>


            
            </form>
                
             
               
</body>
</html>
