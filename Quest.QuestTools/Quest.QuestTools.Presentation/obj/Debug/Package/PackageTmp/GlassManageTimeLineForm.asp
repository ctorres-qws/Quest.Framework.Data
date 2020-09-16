<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Edit Form Page for Glass Items-->
<!-- Submits to page GlassManageCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>

  <style>
	input{
	 text-indent: 100px;
	}
  </style>

<% 
GID = Request.QueryString("Gid")
ticket = Request.QueryString("ticket")

select case ticket
	case "active"
		returnEmail = "GlassToolViewActive.asp"
	case "optima"
		returnEmail = "GlassToolViewOptima.asp"
	case "waiting"
		returnEmail = "GlassToolViewWait.asp"
	case "received"
		returnEmail = "GlassToolViewReceived.asp"
	case "waitingcom"
		returnEmail = "GlassToolViewWaitCommercial.asp"
	case "waitingser"
		returnEmail = "GlassToolViewWaitService.asp"
	case "completed"
		returnEmail = "GlassToolViewCompleted.asp"
	case "shipped"
		returnEmail = "GlassToolViewShipped.asp"
	case Else
		returnEmail = "GlassManage.asp"
	End select

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_GLASSDB"
		rs.Cursortype = GetDBCursorType
		rs.Locktype = GetDBLockType
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & GID

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="<%Response.write returnEmail%>" target="_self">Select Glass</a>
    </div>
    <form id="GlassEdit" title="Edit Glass Timeline" class="panel" action="GlassDelConf.asp" name="GlassEdit"  method="GET" target="_self" selected="true" > 
  <h2><% response.write Trim(rs.fields("ID")) %>: <% response.write Trim(rs.fields("JOB")) %><% response.write Trim(rs.fields("FLOOR")) %>-<% response.write Trim(rs.fields("TAG")) %></h2>
	<h2> Please input all Dates in the Form MM/DD/YYYY --- IMPORTANT for Database Consistency among many users--- </h2>
	<fieldset>

	        <div class="row">
                <label>Req. Date</label>
                <input type="text" name='REQUIREDDATE' id='REQUIREDDATE' size='8' value = '<% response.write Trim(rs.fields("REQUIREDDATE")) %>'>
            </div> 
			<div class="row">
                <label>Input/Order Date</label>
                <input type="text" name='InputDATE' id='InputDATE' size='8' value = '<% response.write Trim(rs.fields("InputDATE")) %>'>
            </div> 
			<div class="row">
                <label>Optima Date</label>
                <input type="text" name='OPTIMADATE' id='OPTIMADATE' size='8' value = '<% response.write Trim(rs.fields("OptimaDATE")) %>'>
            </div> 
			<div class="row">
                <label>Ext Expected Date</label>
                <input type="text" name='EXTEXPECTED' id='EXTEXPECTED' size='8' value = '<% response.write Trim(rs.fields("ExtExpected")) %>'>
            </div> 
			<div class="row">
                <label>Ext Received Date</label>
                <input type="text" name='EXTRECEIVED' id='EXTRECEIVED' size='8' value = '<% response.write Trim(rs.fields("Extreceived")) %>'>
            </div> 
			<div class="row">
                <label>Int Expected Date</label>
                <input type="text" name='INTEXPECTED' id='INTEXPECTED'' size='8' value = '<% response.write Trim(rs.fields("IntExpected")) %>'>
            </div> 
			<div class="row">
                <label>Int Received Date</label>
                <input type="text" name='INTRECEIVED' id='INTRECEIVED' size='8' value = '<% response.write Trim(rs.fields("IntReceived")) %>'>
            </div> 
			<div class="row">
                <label>Completed Date</label>
                <input type="text" name='COMPLETEDDATE' id='COMPLETEDDATE' size='8' value = '<% response.write Trim(rs.fields("COMPLETEDDATE")) %>'>
            </div> 
			<div class="row">
                <label>Ship Date</label>
                <input type="text" name='SHIPDATE' id='SHIPDATE' size='8' value = '<% response.write Trim(rs.fields("SHIPDATE")) %>'>
            </div> 

			</fieldset>

			<h2>Old Timeline</h2>
			<h2> Cardinal</h2>
			<fieldset>
			<div class="row">
                <label>Order Date</label>
                <input type="text" name='CARDINALSENT' id='CARDINALSENT' size='8' value = '<% response.write Trim(rs.fields("CARDINALSENT")) %>'>
            </div> 
			<div class="row">
                <label>Expected Date</label>
                <input type="text" name='CARDINALEXPECTED' id='CARDINALEXPECTED' size='8' value = '<% response.write Trim(rs.fields("CARDINALEXPECTED")) %>'>
            </div> 
			<div class="row">
                <label>Received Date</label>
                <input type="text" name='CARDINALReceived' id='CARDINALReceived' size='8' value = '<% response.write Trim(rs.fields("CARDINALReceived")) %>'>
            </div> 
			</fieldset>
			<h2> Quick Temp </h2>
			<fieldset>
			<div class="row">
                <label>Order Date</label>
                <input type="text" name='QuickTempSent' id='QuickTempSent' size='8' value = '<% response.write Trim(rs.fields("QuickTempSent")) %>'>
            </div> 
            <div class="row">
                <label>Received Date</label>
                <input type="text" name='QuickTempReceived' id='QuickTempReceived' size='8' value = '<% response.write Trim(rs.fields("QuickTempReceived")) %>'>
            </div>

						<input type="hidden" name='GID' id='GID' value="<%response.write GID %>" />
						<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>" />
</fieldset>

        <BR>

		<a class="whiteButton" onClick="GlassEdit.action='GlassManageTimeLineConf.asp'; GlassEdit.submit()">Submit Changes</a><BR>
		<input class="redButton" type='submit' value='Delete Glass' onclick='return confirm(" Delete This Piece of Glass from Inventory? \n Only Delete if certain it will not affect other items. ")'></td>

            </form>

</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

