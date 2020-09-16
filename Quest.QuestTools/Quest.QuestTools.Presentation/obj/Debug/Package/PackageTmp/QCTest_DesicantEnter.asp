                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		<!-- Testing Results stored in the system - Designed for Victor Babuskins - November 2014, Michael Bernholtz-->
<!-- Entry Page - Confirms to QCTest_DesicantConf.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>New Desicant Result</title>
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
        <a class="button leftButton" type="cancel" href="QCTEST_Desicant.asp" target="_self">Desicant Test</a>
        </div>

            
              <form id="enter" title="Enter New Result" class="panel" name="enter" action="QCTEST_Desicantconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Desicant Result:</h2>
  
                       
       <fieldset>

	<div class="row">
		<label>Date</label>
		<%
		currentDate = Date()
			mm = Month(currentDate)
			dd = Day(currentDate)
			yy = Year(currentDate)
			IF len(mm) = 1 THEN
			  mm = "0" & mm
			END IF
			IF len(dd) = 1 THEN
			  dd = "0" & dd
			END IF
			currentDate = yy & "-" & mm & "-" & dd 
		%>
		<input type="Date" name='Date' id='Date' value = '<%response.write currentDate %>' >
	</div>

	<div class="row">
		<label>Time </label>
		<input type="text" name='Time' id='Time' >
    </div>

	<div class="row">
		<label>T1</label>
		<input type="text" name='T1' id='T1' >
    </div>

	<div class="row">
		<label>T2</label>
		<input type="text" name='T2' id='T2' >
    </div>

	<div class="row">
		<label>Result</label>
		<input type="text" name='Result' id='Result' >
    </div>
	
	<div class="row">
		<label>Temp Passed?</label>
        <input type="checkbox" name='Temp' id='Temp' checked>
    </div>  
	
	<div class="row">
		<label>Tested By</label>
		<input type="text" name='Initials' id='Initials' >
    </div>
	
	<div class="row">
		<label>Notes</label>
		<input type="text" name='Notes' id='Notes' >
    </div>

    <a class="whiteButton" href="javascript:enter.submit()" target='_Self'>Submit</a>
            
         
</fieldset>


            
            </form>
           
</body>
</html>
