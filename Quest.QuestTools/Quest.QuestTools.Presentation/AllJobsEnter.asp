<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->		 
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

<% 
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "Select * FROM XQSU_GlassTypes order by DESCRIPTION"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		
		Set rs3 = Server.CreateObject("adodb.recordset")
		strSQL3 = "Select * FROM XQSU_OTSpacer"
		rs3.Cursortype = 2
		rs3.Locktype = 3
		rs3.Open strSQL3, DBConnection
		
%>

 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>
		


              <form id="enter" title="Enter New Job" class="panel" name="enter" action="AllJobsconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Job:</h2>

       <fieldset>

	<div class="row">
		<label>Job Code</label>
		<input type="text" name='JOB' id='JOB' >
	</div>
		
	<div class="row">
		<label>Parent Code</label>
		<input type="text" name='PARENT' id='PARENT' >
	</div>

	<div class="row">
		<label>Job Name </label>
		<input type="text" name='JOB_NAME' id='JOB_NAME' >
    </div>

    <div class="row">
		<label>Job Address </label>
		<input type="text" name='JOB_ADDRESS' id='JOB_ADDRESS' >
    </div>
	<div class="row">
		<label>Job City </label>
		<input type="text" name='JOB_CITY' id='JOB_CITY' >
    </div>
	
	<div class="row">
		<label>Job Country</label>
        <select name= 'JOB_COUNTRY' id = 'JOB_COUNTRY'>
			<option value = "" selected >Please Select</option>
			<option value="CA">Canada</option>
			<option value="US">United States</option>
		</select>
     </div>	

	 <div class="row">
		<label># of Floors</label>
		<input type="number" name='Floors' id='Floors' value = "0">
    </div>
	
	<div class="row">
		<label>(PM)Recipient List</label><br>
		<input type="text" name='RList' id='RList'>
    </div>
	
	<div class="row">
		<label>(PM)Imp Record</label><br>
		<input type="text" name='IMPREC' id='IMPREC'>
    </div>
	
	<div class="row">
		<label>(PM)Imp Address</label><br>
		<input type="text" name='IMPADD' id='IMPADD'>
    </div>
	
	<div class="row">
		<label>(PM)Imp Tax ID</label><br>
		<input type="text" name='IMPTAX' id='IMPTAX'>
    </div>
	
	<div class="row">
		<label>(PM)Exp Tax ID</label><br>
		<input type="text" name='EXPTAX' id='EXPTAX'>
    </div>
	
	<div class="row">
		<label>Ext_Colour</label><br>
		<input type="text" name='EXT_COLOUR' id='EXT_COLOUR' maxlength=250 >
    </div>
	
	<div class="row">
		<label>Int_Colour </label><br>
		<input type="text" name='INT_COLOUR' id='INT_COLOUR' maxlength=250 >
    </div>
	
	<div class="row">
		<label>Manager </label>
		<input type="text" name='Manager' id='Manager' >
    </div>
	
	<div class="row">
		<label>Mgr Email</label>
		<input type="text" name='ManagerEmail' id='ManagerEmail' >
    </div>
	 
	<div class="row">
		<label>Engineer </label>
		<input type="text" name='Engineer' id='Engineer' >
    </div>
	
	<div class="row">
		<label>Eng Email</label>
		<input type="text" name='EngineerEmail' id='EngineerEmail' >
    </div>
	
	<div class="row">
		<label>Material</label>
        <select name= 'MATERIAL' id = 'MATERIAL'>
			<option value = "" selected >Please Select</option>
			<option value="Ecowall">Ecowall</option>
		</select>
     </div>	
	
	<div class="row">
		<label>Ext Glass (GL1)</label>
		<select name='ExtGlass'>
		<option value = "" selected >Please Select</option>
		<%
			rs2.movefirst
			do while not rs2.eof
			Response.Write "<option name='ExtGlass'"
			Response.Write " option value = '"
			Response.Write rs2("Type")
			Response.Write "'>"
			Response.Write rs2("Description")
			rs2.movenext
			loop
		%>			
		</select>
	 </div>
	 
	 <div class="row">
		<label>Int Glass (GL1)</label>
		<select name='IntGlass'>
		<option value = "" selected >Please Select</option>
		<%
			rs2.movefirst
			do while not rs2.eof
			Response.Write "<option name='ExtGlass'"
			Response.Write " option value = '"
			Response.Write rs2("Type")
			Response.Write "'>"
			Response.Write rs2("Description")
			rs2.movenext
			loop
		%>			
		</select>
	 </div>
	 
	 <div class="row">
		<label>Ext Glass Door (GL2)</label>
		<select name='ExtGlassDoor'>
		<option value = "" selected >Please Select</option>
		<%
			rs2.movefirst
			do while not rs2.eof
			Response.Write "<option name='ExtGlass'"
			Response.Write " option value = '"
			Response.Write rs2("Type")
			Response.Write "'>"
			Response.Write rs2("Description")
			rs2.movenext
			loop
		%>			
		</select>
	 </div>
	 
	 <div class="row">
		<label>Int Glass Door (GL2)</label>
		<select name='IntGlassDoor'>
		<option value = "" selected >Please Select</option>	
		<%
			rs2.movefirst
			do while not rs2.eof
			Response.Write "<option name='ExtGlass'"
			Response.Write " option value = '"
			Response.Write rs2("Type")
			Response.Write "'>"
			Response.Write rs2("Description")
			rs2.movenext
			loop
		%>			
		</select>
	 </div>

	
	<div class="row">
		<label>Frame Style</label>
		<select name= 'Fstyle' id = 'Fstyle'>
			<option value = "" selected >Please Select</option>	
			<option value="EW">ECOWALL WITH DRAINAGE</option>
			<option value="EWRS">ECOWALL RECEPTOR SYSTEM B.C.</option>
		</select>
	</div>
	
	<div class="row">
		<label>Mullion under Sliding Door</label>
        <select name= 'SDsill' id = 'SDsill'>
			<option value = "" selected >Please Select</option>
			<option value="No Sill">No Mullion</option>
			<option value="Sill">Mullion </option>
		</select>
    </div>
	<div class="row">
		<label>Mullion under Swing Door</label>
        <select name= 'SWsill' id = 'SWsill'>
			<option value = "" selected >Please Select</option>
			<option value="No Sill">No Mullion</option>
			<option value="Sill">Mullion </option>
		</select>
    </div>
	
	<div class="row">
		<label>Window Glass Thickness (mm)</label>
		<select name='WThickSpacer' id = 'WThickSpacer'>
		<option value = "" selected >Please Select</option>
		<%
			rs3.movefirst
			do while not rs3.eof
			Response.Write "<option name='ThickSpacer'"
			Response.Write " option value = '"
			Response.Write rs3("OTMM") & "|"
			Response.Write rs3("SPACER")
			Response.Write "'>"
			Response.Write rs3("OT") & " - " & rs3("OTMM") 
			rs3.movenext
			loop
		%>			
		</select>
	 </div>	
	<div class="row">
		<label>Swing Door Glass Thickness (mm)</label>
		<select name='SThickSpacer' id = 'SThickSpacer'>
		<option value = "" selected >Please Select</option>
		<%
			rs3.movefirst
			do while not rs3.eof
			Response.Write "<option name='ThickSpacer'"
			Response.Write " option value = '"
			Response.Write rs3("OTMM") & "|"
			Response.Write rs3("SPACER")
			Response.Write "'>"
			Response.Write rs3("OT") & " - " & rs3("OTMM") 
			rs3.movenext
			loop
		%>			
		</select>
	 </div>	
	<div class="row">
		<label>Awning/Casement Glass Thickness (mm)</label>
		<select name='OVThickSpacer' id = 'OVThickSpacer'>
		<option value = "" selected >Please Select</option>
		<%
			rs3.movefirst
			do while not rs3.eof
			Response.Write "<option name='ThickSpacer'"
			Response.Write " option value = '"
			Response.Write rs3("OTMM") & "|"
			Response.Write rs3("SPACER")
			Response.Write "'>"
			Response.Write rs3("OT") & " - " & rs3("OTMM") 
			rs3.movenext
			loop
		%>			
		</select>
	 </div>	
	<div class="row">
		<label>SW Sill Type</label>
        <select name= 'silltype' id = 'silltype'>	
			<option value="" selected >Please Select</option>		
			<option value = "ADAKP">ADA Sill w/ Kickplate </option>
			<option value = "ADA">ADA Sill</option>
			<option value = "FULL">Full Sill</option>
			<option value = "FULLKP">Full Sill w/ Kickplate</option> <!--Unused -->
		</select>
    </div>	
 	
	<div class="row">
		<label>Spacer Colour</label>
        <select name= 'SpacerColour' id = 'spacercolour'>
			<option value="" selected >Please Select</option>		
			<option value="Black">Black</option>
			<option value="Grey">Grey</option>
		</select>
    </div>	
	
	<div class="row">
		<label>Awning Style</label>
        <select name= 'AwnStyle' id = 'Awnstyle'>
			<option value="" selected >Please Select</option>		
			<option value="F92">F92 </option>
			<!--<option value="Metra">Metra </option>-->
			<option value="M90 Outswing">M90 Outswing</option>
			<option value="M90 Inswing">M90 Inswing</option>
		</select>
    </div>	
	
	<div class="row">
		<label>Louver - Glass Stops Needed?</label>
        <select name= 'LouverStyle' id = 'LouverStyle'>
			<option value="" selected >Please Select</option>		
			<option value="Yes">Yes</option>
			<option value="No">No </option>
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
    
	<div class="row">
		<label>Completed</label>
        <input type="checkbox" name='Completed' id='Completed'>
    </div>
	
	<div class="row">
		<label>Max Hoist (in)</label><br>
		<input type="number" name='MaxHoist' id='MaxHoist' value = "0">
    </div>
	
	<div class="row">
		<label>Vertical Stock Length for Order (in)</label><br>
		<input type="number" name='VStockLength' id='VStockLength' value = "0">
    </div>	
	
	<div class="row">
		<label>Job Status</label>
        <select name= 'JobStatus' id = 'JobStatus'>
			<option value = "" selected >Please Select</option>
			<option value="Guaranteed">Guaranteed</option>
			<option value="Measured">Measured</option>
			<option value="Mixed">Mixed</option>
		</select>
    </div>		
	
	<div class="row"> 
		<label>Onsite Date</label>
		<input type="date" name='OnSiteDate' id='OnSiteDate' >	
	</div>	
<!-- ADDITION OF STOP COLOR FOR ORDER ENTRY REQUIREMENTS Dec 13, 2019 ARAMIREZ -->	
<div class="row">
	<label>Stop Color</label>
	<select name= 'StopColor' id = 'StopColor'>
	  <option value="">Please Select</option>
	  <option value="White">White</option>
	  <option value="Color">Color</option>
	</select>
 </div>
		
<!--
	<div class="row">
		<label>Stop Color</label>
		<select name='StopColor'>
		<option value = "" selected >Please Select</option>
		</select>
	 </div>
-->	 
 	<%
		rs2.close
		set rs2 = nothing
		
		rs3.close
		set rs3 = nothing		
	%>	 	

	<div class="row">
		<label>No Doors</label>
        <input type="checkbox" name='NoDoors' id='NoDoors' />  
	</div>	
	
	<div class="row">
		<label>No Awnings</label>
        <input type="checkbox" name='NoAwnings' id='NoAwnings'  />  
	</div>	
	
	<div class="row">
		<label>Screen</label>
        <select name= 'Screen' id = 'Screen'>
			<option value = "" selected >Please Select</option>
			<option value="Yes">Yes</option>
			<option value="No">No</option>
		</select>
    </div>		


                    <a class="whiteButton" href="javascript:enter.submit()" target='_Self'>Submit</a>

</fieldset>

            </form>
	

</body>
</html>
