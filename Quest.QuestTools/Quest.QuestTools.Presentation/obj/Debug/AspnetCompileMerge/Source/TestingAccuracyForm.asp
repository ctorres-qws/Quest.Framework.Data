<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Testing Form to be implemented by Daniel Zalcman -->
<!-- Form collects Testing lengths from all the automated Saws to confirm accuracy within 1/16 -->
<!-- Created May 2016, by Michael Bernholtz for Daniel Zalcman, Confirmed by Jody Cash -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Saw Accuracy Test Form</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
    </head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle">Accuracy</h1>
		<a class="button leftButton" type="cancel" href="index.html#_QCT" target="_self">QCT</a>
    </div>
	<form id="enter" title="Enter New Glass Form" class="panel" name="enter" action="TestingAccuracyConf.asp" method="GET" target="_self" selected="true">
		<h3>Enter New Glass Information:</h3>
		<h3>Two Stage Accuracy Test:</h3>
		<h2> Step 1: No Bar (20)</h2>
		<h2> Step 2: With Bar (60)</h2>
        <fieldset>
		

		<div class="row">
			<label>Date </label>
			<input type="date" name="dateIn" value="<%response.write now %>" />
        </div>
		<div class="row">
			<label>Ameri - 20</label>
			<input type="number" name='Ameri20' id='Ameri20'  value = '20' />
        </div>
		        </div>
		<div class="row">
			<label>Ameri - 60</label>
			<input type="number" name='Ameri60' id='Ameri60'  value = '60' />
        </div>
				<div class="row">
			<label>Adjustment </label>
			<input type="text" name='AmeriAdjust' id='AmeriAdjust' />
        </div>
		<div class="row">
			<label>Material </label>
			<input type="text" name='AmeriMat' id='AmeriMat' />
        </div>
		<div class="row">
			<label>ProLine- 20</label>
			<input type="number" name='Pro20' id='Pro20'  value = '20' />
        </div>
		        </div>
		<div class="row">
			<label>ProLine - 60</label>
			<input type="number" name='Pro60' id='Pro60'  value = '60' />
        </div>
		<div class="row">
			<label>Adjustment </label>
			<input type="text" name='ProAdjust' id='ProAdjust' />
        </div>
				<div class="row">
			<label>Material </label>
			<input type="text" name='ProMat' id='ProMat' />
        </div>
		<div class="row">
			<label>Pertici - 20</label>
			<input type="number" name='Pertici20' id='Pertici20'  value = '20' />
        </div>
		        </div>
		<div class="row">
			<label>Pertici - 60</label>
			<input type="number" name='Pertici60' id='Pertici60'  value = '60' />
        </div>
		<div class="row">
			<label>Adjustment </label>
			<input type="text" name='PerticiAdjust' id='PerticiAdjust' />
        </div>
		<div class="row">
			<label>Material</label>
			<input type="text" name='PerticiMat' id='PerticiMat' />
        </div>
		<div class="row">
			<label>2-Ameri - 20</label>
			<input type="number" name='2Ameri20' id='2Ameri20'  value = '20' />
        </div>
		        </div>
		<div class="row">
			<label>2-Ameri - 60</label>
			<input type="number" name='2Ameri60' id='2Ameri60'  value = '60' />
        </div>
				<div class="row">
			<label>Adjustment </label>
			<input type="text" name='2AmeriAdjust' id='2AmeriAdjust' />
        </div>
		<div class="row">
			<label>Material </label>
			<input type="text" name='2AmeriMat' id='2AmeriMat' />
        </div>
				<div class="row">
			<label>2-Pertici - 20</label>
			<input type="number" name='2Pertici20' id='2Pertici20'  value = '20' />
        </div>
		        </div>
		<div class="row">
			<label>2-Pertici - 60</label>
			<input type="number" name='2Pertici60' id='2Pertici60'  value = '60' />
        </div>
		<div class="row">
			<label>Adjustment </label>
			<input type="text" name='2PerticiAdjust' id='2PerticiAdjust' />
        </div>
		<div class="row">
			<label>Material</label>
			<input type="text" name='2PerticiMat' id='2PerticiMat' />		
		<div class="row">
			<label>Tested by </label>
			<input type="text" name='testby' id='testby' value = "Daniel" />
        </div>
		<div class="row">
			<label>Additional Notes </label>
			<input type="text" name='Notes' id='Notes' />
        </div>




		</fieldset>
		<table> 
		
        <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            </form>
                          
</body>
</html>
