<!--#include file="dbpath.asp"-->                     
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Optimization Log Entry Form Designed for Victor-->
<!-- August 2014, by Michael Bernholtz -->
<!-- Form to Enter Optimization Entry information to be updated at time of cut -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optimization Log</title>
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
       <a class="button leftButton" type="cancel" href="index.html#_GlassP" target="_self">Glass Prod</a>
	   </div>
   
   
   
   
            
              <form id="enter" title="Admin Tools" class="panel" name="enter" action="OptimizationLogConf.asp" method="GET" target="_self" selected="true">
              
                              


        <h2>Add to Optimization Log:</h2>
		<h3>All Fields with * are required</h3>
                       
            <fieldset>


            <div class="row">
                <label>*Job</label>
				<select name="Job">
				<% ActiveOnly = True %>
				<option value="SRV">SRV</option>
				<option value="EB">EB</option> 
                <!--#include file="Jobslist.inc"-->
				</select>
            </div>

            <div class="row">
                <label>*Floor</label>
                <input type="text" name='floor' id='floor' required="required">
            </div>
			
			  <div class="row">
                <label>*Glass</label>
                <select name="Glass">
<% mat = "Glass" %>
<% entertype = "Code" %>
<!--#include file="QSU.inc"-->
</select>
</div>
			
			<div class="row">
                <label>*Type</label>
                <select name="inventorytype">
					<option value="SU">SU</option>
					<option value="OV">OV</option>
					<option value="SP">SP</option>
					<option value="SU/OV">SU/OV</option>
					<option value="Door">Door</option>
					<option value="Other">Other</option>
				</select>
            </div>
            
			<div class="row">
                <label>*Opt File</label>
                <input type="text" name='opfile' id='opfile' required="required">
            </div>
			<div class="row">
                <label>*# of Lites</label>
                <input type="number" name='Lites' id='Lites' value = "1" required="required">
            </div>
			<div class="row">
                <label>Bending File</label>
                <input type="text" name='bendfile' id='bendfile' required="required">
            </div>
			<div class="row">
                <label>*Opt Date</label>
                <input type="text" name='opdate' id='opdate' required="required" value='<% response.write Date %>'>
            </div>
	        

                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>


            
            </form>
                
<%

DBConnection.close
set DBConnection=nothing
%>             
               
</body>
</html>
