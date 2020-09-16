<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 
<!--AllJobs Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->
<!-- Updated for Global Variables Updates by Annabel Ramirez August 2019 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Job Summary</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <link rel="stylesheet" href="/styles/palette-color-picker.css" type="text/css" />
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script src="//code.jquery.com/jquery-1.11.3.min.js"></script>

  <script type="application/x-javascript" src="/js/palette-color-picker.js"></script>
  
 
  


<% 
JID = Request.QueryString("Jid")


		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_Jobs WHERE ID = " & JID
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		
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
'Added for Scan to Print - Jun 2020, CTORRES
		Set rsColors = Server.CreateObject("adodb.recordset")
		strSQLColors = "Select ColorName,ColorHexRGB FROM X_Shipping_Colors"
		rsColors.Cursortype = 2
		rsColors.Locktype = 3
		rsColors.Open strSQLColors, DBConnection		

%>
 <script type="text/javascript">
	iui.animOn = true;

	$(document).ready(function(){
		debugger;
		$('[name="ShippingLabelColorInfo"]').paletteColorPicker({
			colors: JSON.parse($("#ShippingLabelColors").text()),
			//custom_class:'double',
  			position:'upside',
			insert: 'after',
			clear_btn: 'first'
		
	});
		$('[data-color="'+ $("#ShippingLabelColor").val() +'"]').click()
	});
	function onColorSelected(name){
		debugger;
		$("#ShippingLabelColor").val(name.substring(name.length - 6));
	}
  </script>
	</head>
<body >
	<div id="ShippingLabelColors" style="display: none;">
		<%
			Response.write "["
			i = 0
			rsColors.movefirst
			do while not rsColors.eof
				if i <> 0 then Response.write ","
				Response.write "{" & """ColorName"": """ & rsColors("ColorName")&""", ""ColorHexRGB"": """ & rsColors("ColorHexRGB")& """}"
				i = 1
				rsColors.movenext
			loop
			Response.write "]"
		%>
	</div>
		
	</div>
    <div class="toolbar">
        <h1 id="pageTitle">Manage Job Summary</h1>
                <a class="button leftButton" type="cancel" href="AllJobsReport.asp" target="_self">All Jobs</a>

    </div>		
	

    <form id="JobEdit" title="Edit Job Summary" class="panel" action="AllJobsEditConf.asp" name="JobEdit"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	
	<div class="row">
		<label>Job Code </label>
		<input type="text" name='JOB' id='JOB' value ='<% response.write Trim(rs.fields("JOB")) %>' >
	</div>
	
	<div class="row">
		<label>Parent Code </label>
		<input type="text" name='PARENT' id='PARENT' value ='<% response.write Trim(rs.fields("PARENT")) %>' >
	</div>
	<div class="row">
		<label>Job Name </label>
		<input type="text" name='JOB_NAME' id='JOB_NAME' value ='<% response.write Trim(rs.fields("JOB_NAME")) %>' >
    </div>

    <div class="row">
		<label>Job Address </label>
		<input type="text" name='JOB_ADDRESS' id='JOB_ADDRESS' value ='<% response.write Trim(rs.fields("JOB_ADDRESS")) %>'>
    </div>
	<div class="row">
		<label>Job City </label>
		<input type="text" name='JOB_CITY' id='JOB_CITY' value ='<% response.write Trim(rs.fields("JOB_CITY")) %>' >
    </div>
	<div class="row">
		<label>Job Country</label>
        <select name= 'JOB_COUNTRY' id = 'JOB_COUNTRY'>
			<option value="" <% if Trim(rs.fields("JOB_COUNTRY")) = "" then response.write "Selected"%> >Please Select</option>
			<option value="CA" <% if Trim(rs.fields("JOB_COUNTRY")) = "CA" then response.write "Selected"%> >Canada</option>
			<option value="US" <% if Trim(rs.fields("JOB_COUNTRY")) = "US" then response.write "Selected"%> >United States</option>
		</select>
     </div>		
	<div class="row">
		<label># of Floors</label>
		<input type="number" name='Floors' id='Floors' value ='<% response.write Trim(rs.fields("FLOORS")) %>' >
    </div>
	
	<div class="row">
		<label>(PM)Recipient List</label><br>
		<input type="text" name='RList' id='RList' value ='<% response.write Trim(rs.fields("RecipientList")) %>' >
    </div>
	
	<div class="row">
		<label>(PM)Imp Record</label><br>
		<input type="text" name='IMPREC' id='IMPREC' value ='<% response.write Trim(rs.fields("ImporterRecord")) %>' >
    </div>
	
	<div class="row">
		<label>(PM)Imp Address</label><br>
		<input type="text" name='IMPADD' id='IMPADD' value ='<% response.write Trim(rs.fields("ImporterAddress")) %>' >
    </div>
	
	<div class="row">
		<label>(PM)Imp Tax ID</label><br>
		<input type="text" name='IMPTAX' id='IMPTAX' value ='<% response.write Trim(rs.fields("ImporterTaxID")) %>' >
    </div>
	
	<div class="row">
		<label>(PM)Exp Tax ID</label><br>
		<input type="text" name='EXPTAX' id='EXPTAX' value ='<% response.write Trim(rs.fields("ExporterTaxID")) %>' >
    </div>
	
	<div class="row">
		<label>Ext_Colour</label>
		<input type="text" name='EXT_COLOUR' id='EXT_COLOUR' value ='<% response.write Trim(rs.fields("EXT_COLOUR")) %>' >
    </div>

	<div class="row">
		<label>Int_Colour </label>
		<input type="text" name='INT_COLOUR' id='INT_COLOUR' value ='<% response.write Trim(rs.fields("INT_COLOUR")) %>' >
    </div>

	<div class="row">
		<label>Manager </label>
		<input type="text" name='MANAGER' id='MANAGER' value ='<% response.write Trim(rs.fields("MANAGER")) %>' >
    </div>
	
	<div class="row">
		<label>Mgr Email</label>
		<input type="text" name='MANAGEREMAIL' id='MANAGEREMAIL' value ='<% response.write Trim(rs.fields("MANAGEREMAIL")) %>' >
    </div>

	<div class="row">
		<label>Engineer</label>
		<input type="text" name='ENGINEER' id='ENGINEER' value ='<% response.write Trim(rs.fields("ENGINEER")) %>' >
    </div>	
	
	<div class="row">
		<label>Eng Email </label>
		<input type="text" name='ENGINEEREMAIL' id='ENGINEEREMAIL' value ='<% response.write Trim(rs.fields("ENGINEEREMAIL")) %>' >
    </div>	
	
	
	<div class="row">
		<label>Material</label>
        <select name= 'MATERIAL' id = 'MATERIAL'>
			<option value="" <% if Trim(rs.fields("MATERIAL")) = "" then response.write "Selected"%> >Please Select</option>
			<option value="Ecowall" <% if Trim(rs.fields("MATERIAL")) = "Ecowall" then response.write "Selected"%> >Ecowall</option>
			<option value="Q4750" <% if Trim(rs.fields("MATERIAL")) = "Q4750" then response.write "Selected"%> >Q4750</option>
		</select>
     </div>	

	 
	 <div class="row">
		<label>Ext Glass (GL1)</label>
		<select name='ExtGlass'>
		<option value="" <% if Trim(rs.fields("EXTGlass")) = "" then response.write "Selected"%> >Please Select</option>
	 <%
		rs2.movefirst
		do while not rs2.eof
		Response.Write "<option name='ExtGlass'"
		if rs.fields("EXTGlass") = rs2("Type") then
		Response.write " selected "
		end if
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
		<option value="" <% if Trim(rs.fields("IntGlass")) = "" then response.write "Selected"%> >Please Select</option>			
	 <%
		rs2.movefirst
		do while not rs2.eof
		Response.Write "<option name='IntGlass'"
		if rs.fields("IntGlass") = rs2("Type") then
		Response.write " selected "
		end if
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
		<option value="" <% if Trim(rs.fields("EXTGlassDoor")) = "" then response.write "Selected"%> >Please Select</option>			
	 <%
		rs2.movefirst
		do while not rs2.eof
		Response.Write "<option name='ExtGlassDoor'"
		if rs.fields("EXTGlassDoor") = rs2("Type") then
		Response.write " selected "
		end if
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
		<option value="" <% if Trim(rs.fields("IntGlassDoor")) = "" then response.write "Selected"%> >Please Select</option>			
	 <%
		rs2.movefirst
		do while not rs2.eof
		Response.Write "<option name='IntGlassDoor'"
		if rs.fields("IntGlassDoor") = rs2("Type") then
		Response.write " selected "
		end if
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
			<option value="" <% if Trim(rs.fields("FrameStyle")) = "" then response.write "Selected"%> >Please Select</option>		
			<option value="EW" <% if Trim(rs.fields("FrameStyle")) = "EW" then response.write "Selected"%>>ECOWALL WITH DRAINAGE</option>
			<option value="EWRS" <% if Trim(rs.fields("FrameStyle")) = "EWRS" then response.write "Selected"%>>ECOWALL RECEPTOR SYSTEM B.C.</option>
			<option value="Q4750" <% if Trim(rs.fields("FrameStyle")) = "Q4750" then response.write "Selected"%>>Q4750</option>
			<option value="Q4750D" <% if Trim(rs.fields("FrameStyle")) = "Q4750D" then response.write "Selected"%>>Q4750D WITH DRAINAGE</option>
			<option value="EcoNoD" <% if Trim(rs.fields("FrameStyle")) = "EcoNoD" then response.write "Selected"%>> ECOWALL No DRAINAGE</option>
			<option value="PER" <% if Trim(rs.fields("FrameStyle")) = "PER" then response.write "Selected"%>>PEARL - * BETA</option>
		</select>
	</div>

		<div class="row">
		<label>Mullion under Sliding Door</label>
        <select name= 'SDsill' id = 'SDsill'>
			<option value="" <% if Trim(rs.fields("SDsill")) = "" then response.write "Selected"%> >Please Select</option>				
			<option value="No Sill" <% if Trim(rs.fields("SDsill")) = "No Sill" then response.write "Selected"%>>No Mullion</option>
			<option value="Sill" <% if Trim(rs.fields("SDsill")) = "Sill" then response.write "Selected"%>>Mullion </option>
		</select>
     </div>	
		<div class="row">
		<label>Mullion under Swing Door</label>
        <select name= 'SWsill' id = 'SWsill'>
			<option value="" <% if Trim(rs.fields("SWSill")) = "" then response.write "Selected"%> >Please Select</option>				
			<option value="No Sill" <% if Trim(rs.fields("SWSill")) = "No Sill" then response.write "Selected"%>>No Mullion</option>
			<option value="Sill" <% if Trim(rs.fields("SWSill")) = "Sill" then response.write "Selected"%>>Mullion </option>
		</select>
     </div>	
	<div class="row">
		<label>Window Glass Thickness (mm)</label>
		<select name='WThickSpacer' id = 'WThickSpacer'>
		<option value = "" selected >Please Select</option>
		<%
			rs3.movefirst
			do while not rs3.eof
			Response.Write "<option name='WThickSpacer'"
			if isnull(rs.fields("FixWindowThick")) then
				FixWindowThick=0
			else
				FixWindowThick=CInt(rs.fields("FixWindowThick"))
			end if 
			if isnull(rs.fields("SUSPACER")) then
				SUSPACER=0
			else
				SUSPACER=CInt(rs.fields("SUSPACER"))
			end if 
			OTMM= CInt(rs3.fields("OTMM")) 
			SPACER= CInt(rs3.fields("SPACER"))
			if FixWindowThick = OTMM and SUSPACER= SPACER then
				Response.write " selected "
			end if
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
			Response.Write "<option name='SThickSpacer'"
			if isnull(rs.fields("SwingDoorThick")) then
				SwingDoorThick=0
			else 
				SwingDoorThick=CInt(rs.fields("SwingDoorThick"))
			end if 
			if isnull(rs.fields("SWSpacer")) then
				SWSpacer=0
			else 
				SWSpacer=CInt(rs.fields("SWSpacer"))
			end if 	
			OTMM= CInt(rs3.fields("OTMM")) 
			SPACER= CInt(rs3.fields("SPACER"))
			if SwingDoorThick = OTMM and SWSpacer= SPACER then
				Response.write " selected "
			end if
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
			Response.Write "<option name='OVThickSpacer'"
			if isnull(rs.fields("CasAwnThick")) then
				CasAwnThick=0
			else 
				CasAwnThick=CInt(rs.fields("CasAwnThick"))
			end if 	
			if isnull(rs.fields("OVSpacer")) then
				OVSpacer=0
			else 
				OVSpacer=CInt(rs.fields("OVSpacer"))
			end if 	
			OTMM= CInt(rs3.fields("OTMM")) 
			SPACER= CInt(rs3.fields("SPACER"))
			if CasAwnThick = OTMM and OVSpacer= SPACER then
				Response.write " selected "
			end if
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
			<option value="" <% if Trim(rs.fields("SillType")) = "" then response.write "Selected"%> >Please Select</option>		
			<option value = "ADAKP" <% if Trim(rs.fields("SillType")) = "ADAKP" then response.write "Selected"%> >ADA Sill w/ Kickplate </option>
			<option value = "ADA" <% if Trim(rs.fields("SillType")) = "ADA" then response.write "Selected"%> >ADA Sill</option>
			<option value = "FULL" <% if Trim(rs.fields("SillType")) = "FULL" then response.write "Selected"%> >Full Sill</option>
			<option value = "FULLKP" <% if Trim(rs.fields("SillType")) = "FULLKP" then response.write "Selected"%>>Full Sill w/ Kickplate</option> <!--Unused -->
		</select>
     </div>	
	 
	 	 <div class="row">
		<label>Spacer Colour</label>
        <select name= 'SpacerColour' id = 'spacercolour'>
			<option value="" <% if Trim(rs.fields("SpacerColour")) = "" then response.write "Selected"%> >Please Select</option>				
			<option value="Black" <% if Trim(rs.fields("SpacerColour")) = "Black" then response.write "Selected"%> >Black</option>
			<option value="Grey" <% if Trim(rs.fields("SpacerColour")) = "Grey" then response.write "Selected"%> >Grey</option>
		</select>
     </div>	
	 
	 	 <div class="row">
		<label>Awning Style</label>
        <select name= 'AwnStyle' id = 'Awnstyle'>
			<!--<option value="Quest">Quest</option> DISCONTINUED-->
			<option value="" <% if Trim(rs.fields("AwnStyle")) = "" then response.write "Selected"%> >Please Select</option>				
			<option value="F92" <% if Trim(rs.fields("AwnStyle")) = "F92" then response.write "Selected"%>>F92 </option>
			<!--<option value="Metra" <% if Trim(rs.fields("AwnStyle")) = "Metra" then response.write "Selected"%> >Metra </option>-->
			<option value="M90 Outswing" <% if (Trim(rs.fields("AwnStyle")) = "M90 Outswing" or Trim(rs.fields("AwnStyle")) = "Metra") then response.write "Selected"%> >M90 Outswing</option>
			<option value="M90 Inswing" <% if Trim(rs.fields("AwnStyle")) = "M90 Inswing" then response.write "Selected"%> >M90 Inswing</option>
		</select>
     </div>	
	 
	 <div class="row">
		<label>Louver - Glass Stops Needed?</label>
        <select name= 'LouverStyle' id = 'LouverStyle'>
			<option value="" <% if Trim(rs.fields("LouverStyle")) = "" then response.write "Selected"%> >Please Select</option>			
			<option value="Yes" <% if Trim(rs.fields("LouverStyle")) = "Yes" then response.write "Selected"%> >Yes</option>
			<option value="No" <% if Trim(rs.fields("LouverStyle")) = "No" then response.write "Selected"%> >No </option>
		</select>
     </div>	
	 	
	<div class="row">
		<label>H-Bar and Que-168 Color Match (Int and Ext Share Ext Color)</label>
        <select name= 'ColorMatch' id = 'ColorMatch'>
			<option value="" <% if Trim(rs.fields("ColorMatch")) = "" then response.write "Selected"%> >Please Select</option>				
			<option value="Yes" <% if Trim(rs.fields("ColorMatch")) = "Yes" then response.write "Selected"%> >Yes</option>
			<option value="No" <% if Trim(rs.fields("ColorMatch")) = "No" then response.write "Selected"%> >No</option>
		</select>
     </div>
	 
	 <div class="row">
		<label>Door Flush - (Shim Space for ADA Doors)</label>
        <select name= 'GlassFlush' id = 'GlassFlush'>
			<option value="" <%if Trim(rs.fields("GlassFlush")) = "" then Response.write "selected"%>>Please Select</option>
			<option value="Flush" <%if Trim(rs.fields("GlassFlush")) = "Flush" then Response.write "selected"%>>Yes (Flush)</option>
			<option value="3/8" <%if Trim(rs.fields("GlassFlush")) = "3/8" then Response.write "selected"%>>No (Extended 3/8) </option>
			<option value="5/8" <%if Trim(rs.fields("GlassFlush")) = "5/8" then Response.write "selected"%>>No (Extended 5/8) </option>
			<option value="3/4" <%if Trim(rs.fields("GlassFlush")) = "3/4" then Response.write "selected"%>>No (Extended 3/4) </option>
		</select>
     </div>	
	 
	 <div class="row">
		<label>Sill Hook Setting</label>
        <select name= 'BeautyStyle' id = 'BeautyStyle'>
			<option value="" <%if Trim(rs.fields("BeautyStyle")) = "" then Response.write "selected"%>>Please Select</option>		
			<option value="Clip"  <%if Trim(rs.fields("BeautyStyle")) = "Clip" then Response.write "selected"%>>Without Hook (Que-146)</option>
			<option value="Hook"  <%if Trim(rs.fields("BeautyStyle")) = "Hook" then Response.write "selected"%>>With Hook (Que-144 + Que-201) </option>
		</select>
     </div>	
	 	
	<div class="row">
		<label>Flush Panel Style</label>
        <select name= 'PanelPunch' id = 'PanelPunch'>
			<option value="" <%if Trim(rs.fields("PanelPunch")) = "" then Response.write "selected"%>>Please Select</option>				
			<!--<option value="Yes" <% if Trim(rs.fields("PanelPunch")) = "Yes" then response.write "Selected"%> >Yes</option>
			<option value="No" <% if Trim(rs.fields("PanelPunch")) = "No" then response.write "Selected"%>>No </option> -->
			<option value="R3" <% if Trim(rs.fields("PanelPunch")) = "R3" Or  Trim(rs.fields("PanelPunch")) = "Yes" then response.write "Selected"%> >R3</option>
			<option value="Bent" <% if Trim(rs.fields("PanelPunch")) = "Bent" Or  Trim(rs.fields("PanelPunch")) = "No" then response.write "Selected"%>>Bent</option> 
			
		</select>
     </div>	 	
    
	</div>
		<div class="row">
		<label>R3 Vent Size (in)</label><br>
		<input type="number" name='R3VentSize' id='R3VentSize' value ='<% response.write Trim(rs.fields("R3VentSize")) %>' >
    </div>	 

    <div class="row">
		<label>Is the Job Completed? (Remove from Active List)</label>
        <input type="checkbox" name='Completed' id='Completed' <% if rs.fields("COMPLETED") = TRUE THEN response.write "checked" END IF%>>
    </div>     
	
	
		 <div class="row">
		<label>Max Hoist (in)</label><br>
		<input type="number" name='MaxHoist' id='MaxHoist' value ='<% response.write Trim(rs.fields("MaxHoist")) %>'">
    </div>
	
	
	 <div class="row">
		<label>Vertical Stock Length for Order (in)</label><br>
		<input type="number" name='VStockLength' id='VStockLength' value ='<% response.write Trim(rs.fields("VStockLength")) %>'>
    </div>
	
	
	<div class="row">
		<label>Job Status</label>
        <select name= 'JobStatus' id = 'JobStatus'>
			<option value="" <%if Trim(rs.fields("JobStatus")) = "" then Response.write "selected"%>>Please Select</option>		
			<option value="Guaranteed" <% if Trim(rs.fields("JobStatus")) = "Guaranteed" then response.write "Selected"%> >Guaranteed</option>
			<option value="Measured" <% if Trim(rs.fields("JobStatus")) = "Measured" then response.write "Selected"%> >Measured</option>
			<option value="Mixed" <% if Trim(rs.fields("JobStatus")) = "Mixed" then response.write "Selected"%> >Mixed</option>
		</select>
     </div>	


<%
	 InputDate = rs.fields("OnSiteDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
%>
			<div class="row"> 
                <label>Onsite Date</label>
                <input type="date" name='OnSiteDate' id='OnSiteDate' value='<%= DateInput %>'>
            </div>	

<!-- ADDITION OF STOP COLOR FOR ORDER ENTRY REQUIREMENTS DEC 13, 2019 BY ARAMIREZ -->
<div class="row">
	<label>Stop Color</label>
	<select name= 'StopColor' id = 'StopColor'>
	<option value="" <%if Trim(rs.fields("StopColor")) = "" then Response.write "selected"%>>Please Select</option>		
	<option value="White" <% if Trim(rs.fields("StopColor")) = "White" then response.write "Selected"%> >White</option>
	<option value="Color" <% if Trim(rs.fields("StopColor")) = "Color" then response.write "Selected"%> >Color</option>
	</select>
 </div>
			
<!-- COMMENTED SINCE JAYESH WILL NOT POPU
	 <div class="row">
		<label>Stop Color</label>
		<select name='StopColor'>
		<option value="" 	<% 'if Trim(rs.fields("StopColor")) = "" then response.write "Selected"%> >Please Select</option>			
-->	
	 <%
		'JOB=TRIM(rs.fields("JOB"))
		'Set rs4 = Server.CreateObject("adodb.recordset")
		'strSQL4 = "Select code FROM Y_COLOR where JOB = '" & JOB & "'"
		'rs4.Cursortype = 2
		'rs4.Locktype = 3
		'rs4.Open strSQL3, DBConnection	
		'if  not rs4.eof then
		'rs4.movefirst
		'end if 
		'do while not rs4.eof
		'Response.Write "<option name='StopColor'"
		'if rs.fields("StopColor") = rs4("Code") then
		'Response.write " selected "
		'end if
		'Response.Write " option value = '"
		'Response.Write rs4("Code")
		'Response.Write "'>"
		'Response.Write rs4("Code")
		'rs4.movenext
		'loop
		
	%>
<!--		</select>
	 </div>	-->
	<%
	rs2.close
		set rs2 = nothing

	rs3.close
		set rs3 = nothing

	'rs4.close
	'	set rs4 = nothing
		
%>	 
	
	<div class="row">
		<label>No Doors</label>
		<input type="checkbox" name='NoDoors' id='NoDoors' <% if rs.fields("NoDoors") = TRUE THEN response.write "checked" END IF%>>
	</div>	
	
	<div class="row">
		<label>No Awnings</label>
		<input type="checkbox" name='NoAwnings' id='NoAwnings' <% if rs.fields("NoAwnings") = TRUE THEN response.write "checked" END IF%>>
	</div>	
	
	<div class="row">
		<label>Screen</label>
        <select name= 'Screen' id = 'Screen'>
			<option value="" <%if Trim(rs.fields("Screen")) = "" then Response.write "selected"%>>Please Select</option>		
			<option value="Yes" <% if Trim(rs.fields("Screen")) = "Yes" then response.write "Selected"%> >Yes</option>
			<option value="No" <% if Trim(rs.fields("Screen")) = "No" then response.write "Selected"%> >No</option>
		</select>
    </div>		

	<!-- ADDITION OF SHIPPING LABEL COLOR FOR SCAN TO PRINT JUN, 2020 BY CTORRES -->	
	<div class="row" style="display: inline-flex;width: 100%;"> 
		<label>Shipping Label Color</label>
		<input type="text" id="ShippingLabelColor" name="ShippingLabelColor" style="display: none;" value='<%Response.write rs.fields("ShippingLabelColor")%>'>
		<input type="text" id="ShippingLabelColorInfo" name="ShippingLabelColorInfo" style="padding-top: 2rem;padding-bottom: 1.5rem;">
	</div>		
	<input type="hidden" name='JID' id='JID' value="<%response.write JID %>" />
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="JobEdit.action='AllJobsEditConf.asp'; JobEdit.submit()">Submit Changes</a><BR>
		
            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

