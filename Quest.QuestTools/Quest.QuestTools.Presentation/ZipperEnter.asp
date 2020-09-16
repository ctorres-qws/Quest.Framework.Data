                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- Form Created December 31, 2014, Michael Bernholtz at request of Slava Kotek, Lev Bedoev, Jody Cash-->
<!-- Entry Form to add items to Zipper-->
<!-- Zipper will be changed to automatic only-->
<!-- Confirms to Zipper table "Roll_Table" using Zipper.conf -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Zipper / Rolling Extrusion</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Zipper" target="_self">Zipper</a>
        </div>

    <form id="enter" title="New Rolling" class="panel" name="enter" action="zipperConf.asp" method="GET" target="_self" selected="true" >
        <h2>Enter Rolling Request</h2>
  <!-- Enter Job, Floor, Profile, Qty, RequestDate, Length, ShearTest -->
		<fieldset>
  
		<div class="row">
		<label>Job </label>
		<select name="Job">
			<% ActiveOnly = True %>
			<option value = "" selected>-</option>
			<!--#include file="JobsList.inc"-->
			</select>
		</div>
		<div class="row">
			<label>Floor</label>
			<input type="text" name='Floor' id='Floor' />
        </div>  
		<div class="row">
			<label>Quantities (Numbers Only)</label><br><br>
			<table>
				<tr>
				<td><label><center>Mullion</center></label><input type="number" name='MullionQty' id='MullionQty' ></td>
				<td><label><center>Sill</center></label><input type="number" name='SillQty' id='SillQty' ></td>
				<td><label><center>Sash</center></label><input type="number" name='SashQty' id='SashnQty' ></td>
				<td><label><center>Jamb / Header</center></label><input type="number" name='JambQty' id='JambQty' ></td>
				</tr>	
			</table>
        </div>
		<div class="row">
			<label>Length (FT) of whole Extrusion (EX: 216)</label><br><br>
			<table>
				<tr>
				<td><label><center>Mullion</center></label><input type="number" name='MullionLft' id='MullionLft' ></td>
				<td><label><center>Sill</center></label><input type="number" name='SillLft' id='SillLft' ></td>
				<td><label><center>Sash</center></label><input type="number" name='SashLft' id='SashnLft' ></td>
				<td><label><center>Jamb / Header</center></label><input type="number" name='JambLft' id='JambLft' ></td>
				</tr>	
			</table>
        </div> 
         <div class="row">
			<label>ShearTest (Unless Exception)</label>
			<input type="text" name='sheartest' id='sheartest' value="Yes" />
        </div>
 </fieldset>
            
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
            </form>
  <%
DBconnection.close
Set DBConnection = Nothing
%>  
         
               
</body>
</html>
