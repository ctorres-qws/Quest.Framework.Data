<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
	Dim str_QueryPeriod, str_PeriodYear, str_PeriodMonth
	str_QueryPeriod = Request("PeriodYear") & Request("PeriodMonth")
	If str_QueryPeriod  <> "" Then
		str_PeriodYear = Left(str_QueryPeriod, 4)
		str_PeriodMonth = Right(str_QueryPeriod, 2)
	End If
	Dim str_Period, str_PriceSapaMill, str_PriceSapaMontreal, str_PriceExtal, str_PriceMetra, str_PriceCanArt, str_PriceApel
	DBOpen DBConnection, True
	If Request("qwsAction") = "PRICE_UPDATE" Then
		If str_QueryPeriod <> "" Then
			Set rs_Data = Server.CreateObject("adodb.recordset")
			strSQL = "SELECT * FROM _qws_Inv_SupplierPrices WHERE Period = " & str_QueryPeriod & " "
			rs_Data.Cursortype = 2
			rs_Data.Locktype = 3
			rs_Data.Open strSQL, DBConnection
			If rs_Data.EOF Then
				rs_Data.AddNew
			End If
			rs_Data.Fields("Period") = str_QueryPeriod
			rs_Data.Fields("SapaMill_H") = Request("PriceSapaMill_H")
			rs_Data.Fields("SapaMill_S") = Request("PriceSapaMill_S")
			rs_Data.Fields("SapaMontreal_H") = Request("PriceSapaMontreal_H")
			rs_Data.Fields("SapaMontreal_S") = Request("PriceSapaMontreal_S")
			rs_Data.Fields("Extal_H") = Request("PriceExtal_H")
			rs_Data.Fields("Extal_S") = Request("PriceExtal_S")
			rs_Data.Fields("Metra_H") = Request("PriceMetra_H")
			rs_Data.Fields("Metra_S") = Request("PriceMetra_S")
			rs_Data.Fields("CanArt_H") = Request("PriceCanArt_H")
			rs_Data.Fields("CanArt_S") = Request("PriceCanArt_S")
			rs_Data.Fields("Apel_H") = Request("PriceApel_H")
			rs_Data.Fields("Apel_S") = Request("PriceApel_S")
			rs_Data.Fields("Default_H") = Request("PriceDefault_H")
			rs_Data.Fields("Default_S") = Request("PriceDefault_S")
			
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

	</style>
	<script>

		function saveSupplierPrices() {
			prices.action = "InventoryReportPrices.asp";
			prices.qwsAction.value = 'PRICE_UPDATE';
			prices.submit();
		}

		function periodChange() {
			prices.action = "InventoryReportPrices.asp";
			prices.qwsAction.value = 'PERIOD_CHANGE';
			prices.submit();
		}

		function editPrice(str_Year, str_Month) {
			frmEditPrice.PeriodYear.value = str_Year;
			frmEditPrice.PeriodMonth.value = str_Month;
			frmEditPrice.qwsAction.value = 'PERIOD_CHANGE';
			frmEditPrice.submit();
		}

		function refreshSearchYear() {
			frmSearch.SearchYear.value = prices.SearchYear.options[prices.SearchYear.options.selectedIndex].value
			frmSearch.submit();
		}

	</script>
	<style>

	#csTable tr:nth-child(odd){
		background-color: #eaeaea;
		color: #0;
	}

	#csTable tr:nth-child(even){
		background-color: #fff;
		color: #0;
	}

	#csTable td, #csTable th { padding-right: 10px; height: 30px; }

	input { margin-left: 200px !important; width: 300px !important; border: 1px solid rgb(221,221,221) !important; border-radius: 5px; margin-bottom: 5px !important; height: 30px !important;}

	</style>
</head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<div style=""><a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportSelectCostBySupplier.asp" target="_self" style="left: 90px;">Inventory Report</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPrices.asp" target="_self" style="left: 216px;">Supplier Prices</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPaintPrices.asp" target="_self" style="left: 333px;">Paint Pricing</a></div>
	</div>

	<form id="prices" title="Inventory Prices" class="panel" name="prices" action="InventoryReportPrices.asp" method="GET" target="_self" selected="true" autocomplete="off">
		<input name="qwsAction" value="" type="hidden">
		<h2>Enter Aluminium Price and Select a Snapshot of Inventory</h2>
		<fieldset>

		<div class="row">
			<label>Inventory Period</label>
			<select id='PeriodMonth' name="PeriodMonth" onchange="periodChange()">
<%
	If str_PeriodMonth <> "" Then Response.Write("<option value='" & Right("0" & str_PeriodMonth, 2) & "'>" & Right("0" & str_PeriodMonth, 2) & "</option>")
	For i = 1 to 12
		Response.Write("<option value='" & Right("0" & i, 2) & "'>" & Right("0" & i, 2) & "</option>")
	Next
%>
			</select>
			<select id='PeriodYear' name="PeriodYear" onchange="periodChange()" >
<%
	If str_PeriodYear <> "" Then Response.Write("<option value='" & str_PeriodYear & "'>" & str_PeriodYear & "</option>")
	For i = Year(Now) to 2010 Step Size - 1
		Response.Write("<option value='" & i & "'>" & i & "</option>")
	Next
%>
			</select>
<%
	Dim str_Year: str_Year = Request("SearchYear")
	If str_Year = "" Then str_Year = Year(Now)

	If str_QueryPeriod <> "" Then
		str_Year = str_PeriodYear
		Set rs_Data = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM _qws_Inv_SupplierPrices WHERE Period = " & str_QueryPeriod & " "
		rs_Data.Cursortype = GetDBCursorType
		rs_Data.Locktype = GetDBLockType
		rs_Data.Open strSQL, DBConnection
		If Not rs_Data.EOF Then
			str_PriceSapaMill_H = rs_Data.Fields("SapaMill_H")
			str_PriceSapaMill_S = rs_Data.Fields("SapaMill_S")
			str_PriceSapaMontreal_H = rs_Data.Fields("SapaMontreal_H")
			str_PriceSapaMontreal_S = rs_Data.Fields("SapaMontreal_S")
			str_PriceExtal_H = rs_Data.Fields("Extal_H")
			str_PriceExtal_S = rs_Data.Fields("Extal_S")
			str_PriceMetra_H = rs_Data.Fields("Metra_H")
			str_PriceMetra_S = rs_Data.Fields("Metra_S")
			str_PriceCanArt_H = rs_Data.Fields("CanArt_H")
			str_PriceCanArt_S = rs_Data.Fields("CanArt_S")
			str_PriceApel_H = rs_Data.Fields("Apel_H")
			str_PriceApel_S = rs_Data.Fields("Apel_S")
			str_PriceDefault_H = rs_Data.Fields("Default_H")
			str_PriceDefault_S = rs_Data.Fields("Default_S")
		End if
		rs_Data.Close: Set rs_Data = Nothing
	End If

%>
		</div>
<% If str_QueryPeriod <> "" Then %>
		<div class="row">
			<label>Supplier Prices ($) for <%= str_QueryPeriod %></label>
		</div>
		<div class="row">
			<label>Sapa Mill - Hollow</label><input type="number" name='PriceSapaMill_H' id='PriceSapaMill' value ="<%= str_PriceSapaMill_H %>" >
		</div>
		<div class="row">
			<label>Sapa Mill - Solid</label><input type="number" name='PriceSapaMill_S' id='PriceSapaMill' value ="<%= str_PriceSapaMill_S %>" >
		</div>

		<div class="row">
			<label>Sapa Montreal - Hollow</label><input type="number" name='PriceSapaMontreal_H' id='PriceSapaMontreal' value ="<%= str_PriceSapaMontreal_H %>" >
		</div>
		<div class="row">
			<label>Sapa Montreal - Solid</label><input type="number" name='PriceSapaMontreal_S' id='PriceSapaMontreal' value ="<%= str_PriceSapaMontreal_S %>" >
		</div>

		<div class="row">
			<label>Extal - Hollow</label><input type="number" name='PriceExtal_H' id='PriceExtal' value ="<%= str_PriceExtal_H %>" >
		</div>
		<div class="row">
			<label>Extal - Solid</label><input type="number" name='PriceExtal_S' id='PriceExtal' value ="<%= str_PriceExtal_S %>" >
		</div>

		<div class="row">
			<label>Metra - Hollow</label><input type="number" name='PriceMetra_H' id='PriceMetra' value ="<%= str_PriceMetra_H %>" >
		</div>
		<div class="row">
			<label>Metra - Solid</label><input type="number" name='PriceMetra_S' id='PriceMetra' value ="<%= str_PriceMetra_S %>" >
		</div>

		<div class="row">
			<label>Can-Art - Hollow</label><input type="number" name='PriceCanArt_H' id='PriceCanArt' value ="<%= str_PriceCanArt_H %>" >
		</div>
		<div class="row">
			<label>Can-Art - Solid</label><input type="number" name='PriceCanArt_S' id='PriceCanArt' value ="<%= str_PriceCanArt_S %>" >
		</div>

		<div class="row">
			<label>Apel - Hollow</label><input type="number" name='PriceApel_H' id='PriceApel' value ="<%= str_PriceApel_H %>" >
		</div>
		<div class="row">
			<label>Apel - Solid</label><input type="number" name='PriceApel_S' id='PriceApel' value ="<%= str_PriceApel_S %>" >
		</div>

		<div class="row">
			<label>Default - Hollow</label><input type="number" name='PriceDefault_H' id='PriceDefault' value ="<%= str_PriceDefault_H %>" >
		</div>
		<div class="row">
			<label>Default - Solid</label><input type="number" name='PriceDefault_S' id='PriceDefault' value ="<%= str_PriceDefault_S %>" >
		</div>


		<div class="row">
			<a class="whiteButton" href="javascript: saveSupplierPrices();">Save Supplier Prices</a>
		</div>
<% End If %>
	</fieldset>
	<BR>
<div class="toolbar">
<label style="font-weight: bold;">Search Year:&nbsp;</label><select name="SearchYear" onchange="refreshSearchYear()" style="height: 25px;">
<%

	For i = Year(Now) to 2014 Step Size - 1
		Dim str_Selected: str_Selected = ""
		If CStr(i) = str_Year Then str_Selected = " selected "
		Response.Write("<option value='" & i & "' " & str_Selected & ">" & i & "</option>")
	Next

%>
</select>
<table border='1' class='sortable' cellpadding="0" cellspacing="0" width="80%" id="csTable" >
	<tr>
		<th rowspan="2">Period</th>
		<th colspan="2">Sapa Mill</th>
		<th colspan="2">Sapa Montreal</th>
		<th colspan="2">Extal</th>
		<th colspan="2">Metra</th>
		<th colspan="2">CanArt</th>
		<th colspan="2">Apel</th>
		<th colspan="2">Default</th>
		<th rowspan=2">Action</th>
	</tr>

	<tr>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
		<th>Hollow</th>
		<th>Solid</th>
	</tr>

<%

	Set rs_Data = Server.CreateObject("adodb.recordset")
	
	strSQL = "SELECT * FROM _qws_Inv_SupplierPrices WHERE Period > " & str_Year & "00" & " AND Period < " & str_Year & "99" & " ORDER BY Period DESC "
	rs_Data.Cursortype = GetDBCursorType
	rs_Data.Locktype = GetDBLockType
	rs_Data.Open strSQL, DBConnection
	Do While Not rs_Data.EOF
		str_Period = rs_Data.Fields("Period")
		str_PriceSapaMill_H = FormatNumber(rs_Data.Fields("SapaMill_H"),2)
		str_PriceSapaMill_S = FormatNumber(rs_Data.Fields("SapaMill_S"),2)
		str_PriceSapaMontreal_H = FormatNumber(rs_Data.Fields("SapaMontreal_H"),2)
		str_PriceSapaMontreal_S = FormatNumber(rs_Data.Fields("SapaMontreal_S"),2)
		str_PriceExtal_H = FormatNumber(rs_Data.Fields("Extal_H"),2)
		str_PriceExtal_S = FormatNumber(rs_Data.Fields("Extal_S"),2)
		str_PriceMetra_H = FormatNumber(rs_Data.Fields("Metra_H"),2)
		str_PriceMetra_S = FormatNumber(rs_Data.Fields("Metra_S"),2)
		str_PriceCanArt_H = FormatNumber(rs_Data.Fields("CanArt_H"),2)
		str_PriceCanArt_S = FormatNumber(rs_Data.Fields("CanArt_S"),2)
		str_PriceApel_H = FormatNumber(rs_Data.Fields("Apel_H"),2)
		str_PriceApel_S = FormatNumber(rs_Data.Fields("Apel_S"),2)
		str_Default_H = FormatNumber(rs_Data.Fields("Default_H"),2)
		str_Default_S = FormatNumber(rs_Data.Fields("Default_S"),2)
%>
<tr>
	<td style="text-align: center;"><%= str_Period %></td>
	<td style="text-align: right;"><%= str_PriceSapaMill_H %></td>
	<td style="text-align: right;"><%= str_PriceSapaMill_S %></td>
	<td style="text-align: right;"><%= str_PriceSapaMontreal_H %></td>
	<td style="text-align: right;"><%= str_PriceSapaMontreal_S %></td>
	<td style="text-align: right;"><%= str_PriceExtal_H %></td>
	<td style="text-align: right;"><%= str_PriceExtal_S %></td>
	<td style="text-align: right;"><%= str_PriceMetra_H %></td>
	<td style="text-align: right;"><%= str_PriceMetra_S %></td>
	<td style="text-align: right;"><%= str_PriceCanArt_H %></td>
	<td style="text-align: right;"><%= str_PriceCanArt_S %></td>
	<td style="text-align: right;"><%= str_PriceApel_H %></td>
	<td style="text-align: right;"><%= str_PriceApel_S %></td>
	<td style="text-align: right;"><%= str_Default_H %></td>
	<td style="text-align: right;"><%= str_Default_S %></td>
	<td style="text-align: center;"><a href="javascript: void();" onclick="editPrice('<%= Left(str_Period,4) %>','<%= Right(str_Period,2) %>');">Edit</a></td>
</tr>
<%
		rs_Data.MoveNext
	Loop
	rs_Data.Close: Set rs_Data = Nothing
%>
</table>
</div>
<%
	'rs.close
	'set rs=nothing
	DBConnection.close
	set DBConnection = nothing
%>
</form>
<form name="frmEditPrice" action="InventoryReportPrices.asp" method="GET">
	<input type="hidden" name="PeriodYear">
	<input type="hidden" name="PeriodMonth">
	<input type="hidden" name="qwsAction">
</form>
<form name="frmSearch" action="InventoryReportPrices.asp" method="GET">
	<input type="hidden" name="SearchYear">
	<input type="hidden" name="qwsAction">
</form>
</body>
</html>
