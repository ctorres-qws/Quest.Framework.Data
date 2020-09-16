<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Glass Profiles Designed for Jody Cash May 2015-->
<!-- Collect Panel Style Information -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Profiles</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Panel</a>
        </div>
   
   
   
   
            
              <form id="enter" title="Enter New Glass Profile" class="panel" name="enter" action="GlassProfileconf.asp" method="POST" target="_self" selected="true">
              
                              


        <h2>Enter New Glass Profile:</h2>
		
                       
    <fieldset>
		<div class="row">
			<label>Profile Name</label>
			<input type="text" name='NAME' id='NAME' >
        </div>
		
		<div class="row">
			<label>Description</label>
			<input type="text" name='Description' id='Description' >
        </div>
	
		<div class="row">
			<label>Job</label>
			<input type="text" name='JOB' id='JOB' >
        </div>

		<div class="row">
			<label>Material</label>
            <Select name='Material'>			
				<option value="Ecowall">Ecowall</option>
				<option value="Q4750">Q4750</option>
			</Select>
        </div>
		<div class="row">
			<label>Ext Glass</label>
			<input type="text" name='ExtGlass' id='ExtGlass' >
        </div>
				<div class="row">
			<label>Ext Glass Door</label>
			<input type="text" name='ExtGlassDoor' id='ExtGlassDoor' >
        </div>
		<div class="row">
			<label>Int Glass</label>
			<input type="text" name='IntGlass' id='IntGlass' >
        </div>
		<div class="row">
			<label>Int Glass Door</label>
			<input type="text" name='IntGlassDoor' id='IntGlassDoor' >
        </div>
		<div class="row">
			<label>SU Thickness</label>
			<input type="number" name='FixWindowThick' id='FixWindowThick' value = 0>
        </div>
		<div class="row">
			<label>SW Thickness</label>
			<input type="number" name='SwingDoorThick' id='SwingDoorThick' value = 0>
        </div>
		<div class="row">
			<label>AWN Thickness</label>
			<input type="number" name='CasAwnThick' id='CasAwnThick' value = 0>
        </div>
		<div class="row">
			<label>SU Spacer</label>
			<input type="number" name='SUSpacer' id='SUSpacer' value = 0>
        </div>
		<div class="row">
			<label>OV Spacer</label>
			<input type="number" name='OVSpacer' id='OVSpacer' value = 0>
        </div>
		<div class="row">
			<label>Spacer Colour</label>
			<input type="text" name='SpacerColour' id='SpacerColour' >
        </div>
		<div class="row">
			<label>Gas</label>
            <Select name='Gas'>			
				<option value="Argon">Argon</option>
				<option value="Air">Air</option>
			</Select>
        </div>

		<div class="row">
			<label>Sill Type</label>
            <Select name='SillType'>			
				<option value="ADA">ADA</option>
				<option value="Full">Full</option>
			</Select>
        </div>
		
		
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>


            
            </form>
                
<%
DBConnection.close
Set DBConnection = nothing
%>		 
               
</body>
</html>
