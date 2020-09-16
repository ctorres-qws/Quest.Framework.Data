<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Created at Request of Shaun Levy with permission from Jody Cash -->
<!--Input form to Collect Aluminium Price and choose a month to get valuation-->
<!--Date: June 27, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add link to Galvanized Sheet
-->
<%
	Dim str_ReportName
	str_ReportName = Request("ReportName")
	Dim str_PriceSapaMill, str_PriceSapaMontreal, str_PriceExtal, str_PriceMetra, str_PriceCanArt, str_PriceApel
	'str_PriceSapaMill = "3.90": str_PriceSapaMontreal = "3.91": str_PriceExtal = "3.92": str_PriceMetra = "3.93": str_PriceCanArt = "3.94": str_PriceApel = "3.95":

	If Request("qwsAction") = "PRICE_UPDATE" Then
		If str_ReportName <> "" Then
			Set rs_Data = Server.CreateObject("adodb.recordset")
			strSQL = "SELECT * FROM INV_Prices WHERE ReportName = '" & str_ReportName & "' "
			rs_Data.Cursortype = 2
			rs_Data.Locktype = 3
			rs_Data.Open strSQL, DBConnection2
			If rs_Data.EOF Then
				rs_Data.AddNew
			End If
			rs_Data.Fields("ReportName") = str_ReportName
			rs_Data.Fields("SapaMill") = Request("PriceSapaMill")
			rs_Data.Fields("SapaMontreal") = Request("PriceSapaMontreal")
			rs_Data.Fields("Extal") = Request("PriceExtal")
			rs_Data.Fields("Metra") = Request("PriceMetra")
			rs_Data.Fields("CanArt") = Request("PriceCanArt")
			rs_Data.Fields("Apel") = Request("PriceApel")
			rs_Data.Update()
			rs_Data.Close: Set rs_Data = Nothing
		End If
	End If

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Inventory Report</title>
	<meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
	<link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	<link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
	<script type="application/x-javascript" src="/iui/iui.js"></script>
	<script type="text/javascript"> iui.animOn = true;</script>
	<style>
		input { margin-left: 100px !important; width: 300px !important; }
	</style>
	<script>

		function saveSupplierPrices() {
			snapshot.action = "InventoryReportSelectCostBySupplier.asp"
			snapshot.qwsAction.value = 'PRICE_UPDATE';
			snapshot.submit();
		}
		
		function periodChange() {
			snapshot.action = "InventoryReportSelectCostBySupplier.asp"
			snapshot.qwsAction.value = 'PERIOD_CHANGE';
			snapshot.submit();
		}

		function viewExtrusionReport() {
			
			var Country= document.getElementById("Country").value;
			
			
			if (Country == "USA") {
				snapshot.action = "InventoryReportValueCostBySupplier_US.asp";
			} else {
				snapshot.action = "InventoryReportValueCostBySupplier.asp";
			}
	
			snapshot.submit();
		}

		function viewSheetReport() {
			
			snapshot.action = "InventoryReportValueSheetCostBySupplier.aspx";
			snapshot.submit();
		}
		
		function viewGalvanizedReport() {
			
			snapshot.action = "InventoryReportValueGalvanizedCostBySupplier.aspx";
			snapshot.submit();
		}

	</script>
</head>
<body>
  <div class="toolbar">
        <h1 id="pageTitle">Hardware Snapshot Report</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
			Actionsite = "InventoryReportValueCostBySupplier_US.asp"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
			Actionsite = "InventoryReportValueCostBySupplier.asp"
			
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Stock<%response.write HomeSiteSuffix%></a>
    </div>








	<form id="snapshot" title="Inventory Snapshot" class="panel" name="snapshot" action="<%response.write Actionsite%>" method="GET" target="_self" selected="true">
		<input name="qwsAction" value="" type="hidden">
		<h2>Select Inventory Period and Country</h2>
		<fieldset>

		<div class="row">
			<label>Inventory Period</label>
			<select id='reportname' name="reportname" onchange="periodChange()" >
<%

	If str_ReportName <> "" Then Response.Write("<option value='" & str_ReportName & "'>" & str_ReportName & "</option>")

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT ReportName FROM INV_Reports WHERE ReportName LIKE '%Y_INV%' ORDER BY ID DESC"
	rs.Cursortype = GetDBCursorType
	rs.Locktype = GetDBLockType
	rs.Open strSQL, DBConnection2

	Do While Not rs.eof
		If str_ReportName = "" Then str_ReportName = rs.fields("ReportName")
		Response.write "<Option value ='"
		Response.write rs.fields("ReportName")
		Response.write "'>" 
		Response.write rs.fields("ReportName")
		Response.write "</option>"
		rs.movenext
	Loop

	'Set rs_Data = Server.CreateObject("adodb.recordset")
	'strSQL = "SELECT * FROM INV_Prices WHERE ReportName = '" & str_ReportName & "' "
	'rs_Data.Cursortype = GetDBCursorType
	'rs_Data.Locktype = GetDBLockType
	'rs_Data.Open strSQL, DBConnection2
	'If Not rs_Data.EOF Then
	'	str_PriceSapaMill = rs_Data.Fields("SapaMill")
	'	str_PriceSapaMontreal = rs_Data.Fields("SapaMontreal")
	'	str_PriceExtal = rs_Data.Fields("Extal")
	'	str_PriceMetra = rs_Data.Fields("Metra")
	'	str_PriceCanArt = rs_Data.Fields("CanArt")
	'	str_PriceApel = rs_Data.Fields("Apel")
	'End if
	'rs_Data.Close: Set rs_Data = Nothing

%>
			</select>
		</div>
		
			<div class="row">
				<label>Report Country (USA / CANADA)</label>
				
								<select id='Country' name="Country" >
		<% 
		if CountryLocation = "USA" then 
		else
		%>
					<Option value ='CANADA'>CANADA</Option>
		<%
		end if
		%>
					<Option value ='USA'>USA</Option>
				</select>
            </div>
<!--//
		<div class="row">
			<label>Supplier Prices ($) - <%= str_ReportName %></label>
		</div>
		<div class="row">
			<label>Sapa Mill</label><input type="number" name='PriceSapaMill' id='PriceSapaMill' value ="<%= str_PriceSapaMill %>" >
		</div>
		<div class="row">
			<label>Sapa Montreal</label><input type="number" name='PriceSapaMontreal' id='PriceSapaMontreal' value ="<%= str_PriceSapaMontreal %>" >
		</div>
		<div class="row">
			<label>Extal</label><input type="number" name='PriceExtal' id='PriceExtal' value ="<%= str_PriceExtal %>" >
		</div>
		<div class="row">
			<label>Metra</label><input type="number" name='PriceMetra' id='PriceMetra' value ="<%= str_PriceMetra %>" >
		</div>
		<div class="row">
			<label>Can-Art</label><input type="number" name='PriceCanArt' id='PriceCanArt' value ="<%= str_PriceCanArt %>" >
		</div>
		<div class="row">
			<label>Apel</label><input type="number" name='PriceApel' id='PriceApel' value ="<%= str_PriceApel %>" >
		</div>
		<div class="row">
			<a class="whiteButton" href="javascript: saveSupplierPrices();">Save Supplier Prices</a>
		</div>
//-->
	</fieldset>
	<BR>
	<a class="whiteButton" href="InventoryReportPrices.asp" target="_self">Supplier Prices</a>
	<a class="whiteButton" href="javascript:viewExtrusionReport()">Extrusion Submit</a>
	<a class="whiteButton" href="javascript: viewSheetReport();">Sheet Submit</a>
	<a class="whiteButton" href="javascript: viewGalvanizedReport();">Galvanized Sheet Submit</a>
	<a class="redButton" href="inventoryReportCreate.asp">Create New Inventory Snapshot</a>
</form>

<%
	rs.close
	set rs=nothing
	DBConnection.close
	set DBConnection = nothing
	DBConnection2.close
	set DBConnection2 = nothing
%>

</body>
</html>
