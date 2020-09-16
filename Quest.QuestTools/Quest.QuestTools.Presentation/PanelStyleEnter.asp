<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Panel Styles Originally Designed for Jody Cash May 2015-->
<!-- Collect Panel Style Information -->
<!-- Brought into Production July 2018 for R3 Panel Processing -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Styles</title>
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
        <a class="button leftButton" type="cancel" href="PanelStyleByJob1.asp" target="_self">Styles</a>
    </div>
   
	<form id="enter" title="Enter New Panel Style" class="panel" name="enter" action="PanelStyleconf.asp" target="_self" selected="true">
		<h2>Enter New Panel Style:</h2>
		
		<fieldset>
			<div class="row">
				<label>Parent Job</label>
				<input type="text" name='Parent' id='Parent' >
			</div>
			
			<div class="row">
				<label>Colour Code</label>
				<input type="text" name='ColorCode' id='ColorCode' >
			</div>

			<div class="row">
				<label>Style Name</label>
				<input type="text" name='NAME' id='NAME' >
			</div>
			
			<div class="row">
				<label>Description</label>
				<input type="text" name='Description' id='Description' >
			</div>
			<div class="row">
				<label>Ext / Int</label>
				<Select name='Side'>
					<option value="Ext.">Exterior</option>
					<option value="Int.">Interior</option>
				</Select>
			</div>
			<div class="row">
				<label>Material</label>
				<Select name='Material'>
				
				<!-- .050" .080" .125" -->
				
					<option value="0.050 INCH ALUM">0.050'' ALUM</option>
					<option value="0.080 INCH ALUM">0.080'' ALUM</option>
					<option value="0.125 INCH ALUM">0.125'' ALUM</option>
					<option value="Steel">Steel</option>
				</Select>
			</div>
			<div class="row">
				<label>Colour</label>
				<input type="text" name='Colour' id='Colour' >
			</div>
			
			<div class="row">
				<label>Notes</label>
				<input type="text" name='Notes' id='Notes' >
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
