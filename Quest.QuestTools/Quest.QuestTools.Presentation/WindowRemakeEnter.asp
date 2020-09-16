                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Window Remake List Entry Form-->
<!--When new glass is called in, this form will be used by all so there is central storage of all information and no chasing -->
<!-- Sends to WindowRemakeConf -->
<!-- Designed August 2014, by Michael Bernholtz at request of Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>New Window for Remake</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Shipping</a>
        </div>

            
              <form id="enter" title="Enter New Job" class="panel" name="enter" action="WindowRemakeconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Remake:</h2>
  
                       
       <fieldset>

	<div class="row">
		<label>Job </label>
		<input type="text" name='JOB' id='JOB' >
	</div>

	<div class="row">
		<label>Floor </label>
		<input type="text" name='Floor' id='Floor' >
	</div>
	<div class="row">
		<label>Tag</label>
		<input type="text" name='Tag' id='Tag' >
	</div>
	<div class="row">
        <label>Break Date</label>
		<input type="text" name='BREAKDATE' id='BREAKDATE' size='8' value='<% response.write Date() %>' >
	</div>  
	<div class="row">
		<label>Break Cause</label>
		<input type="text" name='BreakCause' id='BreakCause' >
	</div>	
	<div class="row">
        <label>Required Date</label>
				<% 
				tenDay = DateAdd("d",10,Date()) 
				%>
		<input type="text" name='RequiredDATE' id='RequiredDATE' size='8' value='<% response.write tenDay %>' >
	</div>  
	<div class="row">
        <label>Send to </label>
		<input type="text" name='Sendto' id='Sendto'  >
	</div>  
    <div class="row">
		<label>Ready</label>
        <input type="checkbox" name='Ready' id='Ready'>
    </div>     
	<div class="row">
        <label>Notes </label>
		<input type="text" name='Notes' id='Notes'  >
	</div>  
            
                    <a class="whiteButton" href="javascript:enter.submit()" target='_Self'>Submit</a>
            
         
</fieldset>


            
            </form>
                
             
               
</body>
</html>
