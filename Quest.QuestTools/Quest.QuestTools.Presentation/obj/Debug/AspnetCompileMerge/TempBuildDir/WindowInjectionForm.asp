<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Window Injection Tool</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
Passkey = "JODY"
SPassword = Session("password")
Password = UCASE(TRIM(Request.Form("pwd")))
%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>

<%
if Password = Passkey or SPassword = Passkey then
Session("password") = Passkey
Session.Timeout=5
%>

        <form id="enter" title="Add Window" class="panel" name="enter" action="WindowInjectionConf.asp" method="GET" target="_self" selected="true">

        <h2>Add Window Manually</h2>
		<fieldset>
			<div class="row">     
                <label>Job</label>
                <input type="text" name='JOB' id='JOB' >
            </div>
			<div class="row">     
                <label>Floor</label>
                <input type="text" name='FLOOR' id='FLOOR' >
            </div>
			<div class="row">     
                <label>TAG</label>
                <input type="text" name='TAG' id='TAG' >
            </div>
			<div class="row">     
                <label>Department</label>
                <select name='DEPT' id='DEPT'>
				<Option value='ASSEMBLY'>Assembly</option>
				<Option value='GLAZING'>Glazing</option>
				<Option value='GLAZING2'>Glazing 2</option>
				</select>
            </div>
			<div class="row">     
                <label>Employee ID</label>
                <input type="text" name='EMPLOYEE' id='EMPLOYEE' >
            </div>
			<div class="row">     
                <label>DD/MM/YYYY</label>
                <input type="text" name='INJECTDATE' id='INJECTDATE' >
            </div>
            
    
             <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
			</fieldset>
        </form>
<%
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="WindowInjectionForm.asp" method="post" target="_self" selected="True">

<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
end if
%>

</body>
</html>
