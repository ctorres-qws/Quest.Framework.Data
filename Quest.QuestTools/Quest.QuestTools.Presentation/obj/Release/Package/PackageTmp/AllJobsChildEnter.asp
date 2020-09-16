<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Add new item to job list-->
<!--AllJobs Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->
<!-- Updated for Global Variables Updates by Annabel Ramirez August 2019 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>New Job</title>
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

      <form id="enter" title="Enter New Job" class="panel" name="enter" action="AllJobsChildconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter Child Job Code for an Existing Parent Job:</h2>

       <fieldset>

	<div class="row">
		<label>Job Code </label>
		<input type="text" name='JOB' id='JOB' >
	</div>
		
	<div class="row">
		<label>Please select the Parent Color</label>
        <select name= 'Parent' id = 'Parent'>
		<option value = "" selected >Please Select</option>	
		<% 
	
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_Jobs WHERE JOB = PARENT Order by JOB ASC"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection
	
	Do while not rs.eof
		Response.write "<option value='" & RS("JOB") & "'>" & RS("JOB") & "</option>"
	rs.movenext
	Loop
	
	rs.close
	set rs = nothing
	DBConnection.close
	Set DBConnection = nothing
	
	%>	

		</select>
     </div>
	 <div class="row">
		<label># of Floors</label>
		<input type="number" name='Floors' id='Floors' value = "0">
    </div>

	
	 
	
	
	<div class="row">
		<label>Ext_Colour</label>
		<input type="text" name='EXT_COLOUR' id='EXT_COLOUR' maxlength=250 >
    </div>
	
	<div class="row">
		<label>Int_Colour </label>
		<input type="text" name='INT_COLOUR' id='INT_COLOUR' maxlength=250 >
    </div>
	
	<div class="row">
		<label>Mullion under Sliding door</label>
        <select name= 'SDsill' id = 'SDsill'>
			<option value = "" selected >Please Select</option>
			<option value="No Sill">No Mullion</option>
			<option value="Sill">Mullion </option>
		</select>
     </div>
	<div class="row">
		<label>Mullion under Swing door</label>
        <select name= 'SWsill' id = 'SWsill'>
			<option value = "" selected >Please Select</option>
			<option value="No Sill">No Mullion</option>
			<option value="Sill">Mullion </option>
		</select>
     </div>

	<div class="row">
		<label>H-Bar and Que-168 Color Match (Int and Ext Share Ext Color)</label>
        <select name= 'ColorMatch' id = 'ColorMatch'>
			<option value = "" selected >Please Select</option>
			<option value="No">No</option>
			<option value="Yes">Yes</option>
		</select>
     </div>
	 
	<div class="row">
		<label>Door Flush - (Shim Space for ADA Doors)</label>
        <select name= 'GlassFlush' id = 'GlassFlush'>
			<option value = "" selected >Please Select</option>
			<option value="Flush">Yes (Flush)</option>
			<option value="3/8">No (Extended 3/8) </option>
			<option value="5/8">No (Extended 5/8) </option>
			<option value="3/4">No (Extended 3/4) </option>
		</select>
     </div>	

	<div class="row">
		<label>Sill Hook Setting</label>
        <select name= 'BeautyStyle' id = 'BeautyStyle'>
			<option value = "" selected >Please Select</option>
			<option value="Clip">Without Hook (Que-146)</option>
			<option value="Hook">With Hook (Que-144 + Que-201) </option>

		</select>
     </div>	 
	
	<div class="row">
		<label>Flush Panel Style </label>
        <select name= 'PanelPunch' id = 'PanelPunch'>
			<option value = "" selected >Please Select</option>
			<option value="R3">R3</option>
			<option value="Bent">Bent</option>
			<!--<option value="Yes">Yes</option>
			<option value="No">No</option>-->
		</select>
    </div>
	
	<div class="row">
		<label>R3VentSize (in)</label><br>
		<input type="number" name='R3VentSize' id='R3VentSize' value = "2">
    </div>

    <a class="whiteButton" href="javascript:enter.submit()" target='_Self'>Submit</a>
	<h2>
		<b>Note: <br>Additional Global Settings have been copied from the Parent Job.<br>If any settings are different, please Edit them now.</b>
	</h2>
</fieldset>

            </form>
		

</body>
</html>
