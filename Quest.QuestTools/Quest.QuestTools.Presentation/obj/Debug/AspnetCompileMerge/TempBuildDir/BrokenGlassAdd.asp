<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 11th, by Michael Bernholtz - Add Page: Marks new Item as broken-->
<!-- Form created at Request of Ariel Aziza Implemented by Michael Bernholtz--> 
<!-- Using Tables: X_Broken -->
<!-- Inputs to BrokenGlassConf.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Broken Glass</title>
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

<body >

    <div class="toolbar">
        <h1 id="pageTitle">Add New Broken Glass</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
	</div>
   
            
    <form id="enter" title="Add New Broken Glass" class="panel" name="enter" action="BrokenGlassConf.asp" method="POST" target="_self" selected="true">
              
		<h2>Fill in Details about Broken Glass</h2>
			  
        <fieldset>               

			<div class="row" >
				<label>Job</label>
				<input type="text" name='job' id='job' >
			</div>
		
			<div class="row" >
				<label>Floor</label>
				<input type="text" name='floor' id='floor'>
			</div>
            		
			<div class="row" >
				<label>Tag</label>
				<input type="text" name='tag' id='tag'>
			</div>
            
			<div class="row" >
				<label>Opening</label>
				<input type="text" name='opening' id='opening'>
			</div>
			
			<div class="row" >
				<label>Width</label>
				<input type="number" name='width' id='width' value = 0>
			</div>
			
			<div class="row" >
				<label>Height</label>
				<input type="number" name='height' id='height' value = 0>
			</div>
			
			<div class="row" >
				<label>Added By</label>
				<input type="text" name='addby' id='addby'>
			</div>
			
			<div class="row" title="What Caused the Break?">
				<label>Reason</label>
				<input type="text" name='reason' id='reason'>
			</div>
			
			<div class="row" >
				<label>Notes</label>
				<input type="text" name='notes' id='notes'>
			</div>
			
			
			<a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
		</fieldset>


            
    </form>

<%

DBConnection.close
set DBConnection = nothing
%>

            
</body>
</html>
