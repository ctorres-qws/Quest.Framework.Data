<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Date: October 1, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add search for ALL and to add period on the add dialog box.			
-->
<%
	Dim str_QueryPeriod, str_PeriodYear, str_PeriodMonth
	Dim dt_Now: dt_Now = Now

	str_QueryPeriod = Request("PeriodYear") & Request("PeriodMonth")
	str_QueryYear = Request("PeriodYear")
	str_QueryMonth = Request("PeriodMonth")
	
	If Request("PeriodYear") <> "ALL" Then
		If str_QueryPeriod  <> "" Then
			str_PeriodYear = Left(str_QueryPeriod, 4)
			str_PeriodMonth = Right(str_QueryPeriod, 2)
			str_QueryPeriod = str_PeriodYear & str_PeriodMonth
		Else
			If Request("Period") <> "" Then
				str_QueryPeriod = Request("Period")
				str_PeriodYear = Left(str_QueryPeriod, 4)
				str_PeriodMonth = Right(str_QueryPeriod, 2)
			Else
				str_PeriodYear = Year(dt_Now)
				str_PeriodMonth = Right("0" & Month(dt_Now),2)
				str_QueryPeriod = str_PeriodYear & str_PeriodMonth
			End If
		End If			
	End If

	Dim str_Period, str_PriceSapaMill, str_PriceSapaMontreal, str_PriceExtal, str_PriceMetra, str_PriceCanArt, str_PriceApel
	DBOpen DBConnection, True

	If Request("qwsAction") = "EDIT" Then
		'Response.Write("UPDATE _qws_Inv_PaintFamily SET PaintFamily='"& Request("FamilyName") & "',[Price]="& Request("Price") & " WHERE ID = " & Request("ID"))
		DBConnection.Execute("UPDATE _qws_Inv_PaintFamily SET PaintFamily='" & Request("FamilyName") & "',Price="& Request("Price") & " WHERE ID = " & Request("ID"))
	End If

	If Request("qwsAction") = "ADD" Then
		Set rsFound = DBConnection.Execute("SELECT * FROM _qws_Inv_PaintFamily WHERE PaintFamily = '" & Request("FamilyName") & "' AND Period = " & Request("PeriodYear") & Request("PeriodMonth"))
		If rsFound.eof Then	
			'Response.Write("INSERT INTO _qws_Inv_PaintFamily (PaintFamily,[Price],[Period]) VALUES('" & Request("FamilyName") & "'," & Request("Price") & "," & Request("Period") & ")")
			DBConnection.Execute("INSERT INTO _qws_Inv_PaintFamily (PaintFamily,[Price],[Period]) VALUES('" & Request("FamilyName") & "'," & Request("Price") & "," & Request("PeriodYear") & Request("PeriodMonth") &  ")")
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
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<style>

	</style>
	<script>
var dialog
$( document ).ready(function() {

	$("#add-price").click(function() {
		frmPaintPrice.qwsAction.value = 'ADD';
		frmPaintPrice.ID.value = '';
		frmPaintPrice.FamilyName.value = '';
		frmPaintPrice.Price.value = '0';

		$(".jqErrMsg").html('');

		dialog.dialog( "open" );
		$("#dialog-form").css("height", "250");
	});

	dialog = $( "#dialog-form" ).dialog({
		autoOpen: false,
		height: 400,
		width: 650,
		modal: true,
		buttons: {
			"Save": savePrice,
			Cancel: function() {
				dialog.dialog( "close" );
			}
		},
		close: function() {
			//form[ 0 ].reset();
			//allFields.removeClass( "ui-state-error" );
		}
	});

		function savePrice() {
			frmPaintPrice.submit();
			dialog.dialog( "close" );
		}

});

var dialogedit
$( document ).ready(function() {

	dialogedit = $( "#dialog-form-edit" ).dialog({
		autoOpen: false,
		height: 400,
		width: 650,
		modal: true,
		buttons: {
			"Save": savePrice,
			Cancel: function() {
				dialogedit.dialog( "close" );
			}
		},
		close: function() {
			//form[ 0 ].reset();
			//allFields.removeClass( "ui-state-error" );
		}
	});

		function savePrice() {
			frmPaintPriceEdit.submit();
			dialogedit.dialog( "close" );
		}

});

	function showEdit(str_ID, str_Name, str_Price) {
		$(".jqErrMsg").html('');
		frmPaintPriceEdit.qwsAction.value = 'EDIT';
		frmPaintPriceEdit.ID.value = str_ID;
		frmPaintPriceEdit.FamilyName.value = str_Name;
		frmPaintPriceEdit.Price.value = str_Price;
		dialogedit.dialog( "open" );
		$("#dialog-form-edit").css("height", "200");
	}

	function showSearch() {
		prices.submit();
	}

	function periodChange(select_item) {
		var periodValue = select_item.value
			 if (periodValue == "ALL") {
				document.getElementById('PeriodMonth').style.display = 'none';
				document.getElementById('PeriodMonth').value = '';
			} 
			else if ($.isNumeric(periodValue)){
				document.getElementById('PeriodMonth').style.display = 'inline-block';
				//Form.fileURL.focus();
			}
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

	//input { margin-left: 200px !important; width: 300px !important; border: 1px solid rgb(221,221,221) !important; border-radius: 5px; margin-bottom: 5px !important; height: 30px !important;}

	.ui-dialog .ui-dialog-content { overflow: hidden !important; }

	fieldset { border: 1px solid rgb(221,221,221); border-radius: 5px; }

	</style>
</head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<div style=""><a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportSelectCostBySupplier.asp" target="_self" style="left: 90px;">Inventory Report</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPrices.asp" target="_self" style="left: 216px;">Supplier Prices</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPaintPrices.asp" target="_self" style="left: 333px;">Paint Pricing</a></div>
		<div style="left: 400px;"><a class="button rightButton" id="add-price" type="cancel" href="#" target="_self" style=""> + Add Paint Price</a> </div>
	</div>

	<form id="prices" title="Inventory Prices" class="panel" name="prices" action="InventoryReportPaintPrices.asp" method="GET" target="_self" selected="true" autocomplete="off">
		<input name="qwsAction" value="" type="hidden">
		<h2>Enter Paint Prices</h2>
		<fieldset>

		<div class="row">
			<label>Inventory Period</label>
			<div style="float: right; padding-right: 80px; padding-top: 12px; ">

			<select id='PeriodYear' name="PeriodYear" onchange="periodChange(this)" >
<%
	Response.Write "<option value='ALL' "
	If Trim(str_QueryYear) = "ALL" Then 
		Response.Write "Selected"
	End If
	Response.Write ">ALL</option>"
	If str_QueryYear <> "" AND Trim(str_QueryYear) <> "ALL" Then 
		Response.Write "<option value='" &str_QueryYear & "' selected>" & str_QueryYear & "</option>"
	End If
	For i = Year(Now) to 2010 Step Size - 1
		Response.Write "<option value='" & i & "'>" & i & "</option>"
	Next		
%>
			</select>
<%			
If str_QueryMonth <> "" And str_QueryYear <> "ALL" Then
%>
			<select id='PeriodMonth' name="PeriodMonth" class="group" style="display:inline-block" onchange="periodChange(this)">
<%
Else
%>
			<select id='PeriodMonth' name="PeriodMonth" class="group" style="display:none">			
<%
End If
	If str_PeriodMonth <> "" And str_QueryYear <> "ALL" Then Response.Write("<option value='" & Right("0" & str_PeriodMonth, 2) & "'>" & Right("0" & str_PeriodMonth, 2) & "</option>")
	For i = 1 to 12
		Response.Write("<option value='" & Right("0" & i, 2) & "'>" & Right("0" & i, 2) & "</option>")
	Next
%>
			</select>
		</div>
&nbsp;<div style=""><a class="button rightButton" href="javascript: void(0)" onclick="showSearch();">Search</a></div>
		</div>
	</fieldset>
	<BR>
<!--//<a class="button leftButton" id="xadd-price" href="#" xtarget="_self" style="left: 500px;">Add</a>//-->
<div class="toolbar">

<table border='1' class='sortable' cellpadding="0" cellspacing="0" width="80%" id="csTable" >
	<tr>
		<th>ID</th>
		<th>Period</th>
		<th style='text-align: left;'>&nbsp;Paint Family</th>
		<th style='text-align: right;'>Price&nbsp;&nbsp;&nbsp;</th>
		<th>Action</th>		
	</tr>

<%

Set rs_Data = Server.CreateObject("adodb.recordset")
If LEFT(str_QueryPeriod,3) = "ALL" then	
	strSQL = "SELECT * FROM _qws_Inv_PaintFamily ORDER BY Period, PaintFamily DESC "
ElseIf str_QueryPeriod <> "" Then
	strSQL = "SELECT * FROM _qws_Inv_PaintFamily WHERE Period = " & str_QueryPeriod & " ORDER BY Period DESC "
End If
	rs_Data.Cursortype = GetDBCursorType
	rs_Data.Locktype = GetDBLockType
	rs_Data.Open strSQL, DBConnection
	Do While Not rs_Data.EOF
		str_PriceID = rs_Data.Fields("ID")
		str_PricePaintFamily = rs_Data.Fields("PaintFamily")
		str_PricePrice = FormatNumber(rs_Data.Fields("Price"),2)
		str_PricePeriod = rs_Data.Fields("Period")

%>
<tr>
	<td style="text-align: center;"><%= str_PriceID %></td>
	<td style="text-align: center;"><%= str_PricePeriod %></td>
	<td style="text-align: left;">&nbsp;<%= str_PricePaintFamily %></td>
	<td style="text-align: right;"><%= str_PricePrice %>&nbsp;</td>
	<td style="text-align: center;"><a href="javascript: void(0);" onClick="showEdit('<%= str_PriceID %>','<%= str_PricePaintFamily %>','<%= str_PricePrice %>');">Edit</a></td>
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

<div id="dialog-form" title="Paint Price">
  <p class="validateTips">All form fields are required.</p>
 
  <form name="frmPaintPrice">
    <fieldset>
	<div class="text ui-widget-content ui-corner-all" style="padding: 6px 6px 6px;>
		<label for="period" style = "">Period  [YYYY] [MM]</label>
		<div style="float: right; padding-right: 10px; padding-top: 0px; padding-bottom: 12px; ">
			<select id='PeriodYear' name="PeriodYear">
<%
	If str_PeriodYear <> "" Then Response.Write "<option value='" & str_PeriodYear & "'>" & str_PeriodYear & "</option>"	
	For i = Year(Now) to 2010 Step Size - 1
		Response.Write "<option value='" & i & "'>" & i & "</option>"
	Next
%>
			</select>
			<select id='PeriodMonth' name="PeriodMonth">
<%
	If str_PeriodMonth <> "" Then Response.Write "<option value='" & Right("0" & str_PeriodMonth, 2) & "'>" & Right("0" & str_PeriodMonth, 2) & "</option>"
	For i = 1 to 12
		Response.Write "<option value='" & Right("0" & i, 2) & "'>" & Right("0" & i, 2) & "</option>"
	Next
%>
			</select>
		</div>
	</div>	
	<div style = "padding-top: 12px;">		
      <label for="name">Paint Family</label>
      <input type="hidden" name="qwsAction" value="">
      <input type="hidden" name="ID" value="">
      <input type="hidden" name="xFamilyName" id="xFamilyName" value="" class="text ui-widget-content ui-corner-all">
<select name="FamilyName" id="FamilyName" class="text ui-widget-content ui-corner-all" style="width: 100%; padding: 6px 6px 6px 8em;">
				<option value="">Select Paint Family</option>
				<option value="PPG Acrynar">PPG Acrynar</option>
				<option value="PPG Duranar">PPG Duranar</option>
				<option value="PPG Duracron">PPG Duracron</option>
				<option value="PPG Duracron White">PPG Duracron White K-1285</option>
				<option value="PPG Duracron Any">PPG Duracron Any</option>
				<option value="PPG Duranar XL">PPG Duranar XL</option>
				<option value="PPG Duranar XL + Basecoat">PPG Duranar XL + Basecoat</option>
				<option value="VALSPAR Acrodize">VALSPAR Acrodize</option>
				<option value="VALSPAR Acroflur">VALSPAR Acroflur</option>
				<option value="VALSPAR Clear Anodize">VALSPAR Clear Anodize</option>
				<option value="VALSPAR Fluropon">VALSPAR Fluropon</option>
				<option value="VALSPAR Fluropon Classic">VALSPAR Fluropon Classic</option>
				<option value="VALSPAR Flurospar">VALSPAR Flurospar</option>
				<option value="VALSPAR Polylure">VALSPAR Polylure</option>
				<option value="Other">Other</option>
</select>	  
      <label for="price">Price</label>
      <input type="text" name="Price" id="Price" value="" class="text ui-widget-content ui-corner-all">
	  </div>
      <!-- Allow form submission with keyboard without duplicating the dialog button -->
      <input type="submit" tabindex="-1" style="position:absolute; top:-1000px">
      <div style="color: #ff0000 !important;" id="jqErrMsg" class="jqErrMsg"><br/>Test Message Test Message Test Message Test Message</div>
    </fieldset>
  </form>
</div>

<div id="dialog-form-edit" title="Paint Price">
  <p class="validateTips">All form fields are required.</p>
 
  <form name="frmPaintPriceEdit">
    <fieldset>
      <label for="name">Paint Family</label>
      <input type="hidden" name="qwsAction" value="">
      <input type="hidden" name="ID" value="">
      <input type="hidden" name="Period" value="<%= str_QueryPeriod %>">
      <input type="hidden" name="xFamilyName" id="xFamilyName" value="" class="text ui-widget-content ui-corner-all">
	  <input type="text" name="FamilyName" id="FamilyName" value="" class="text ui-widget-content ui-corner-all" readonly>
      <label for="price">Price</label>
      <input type="text" name="Price" id="Price" value="" class="text ui-widget-content ui-corner-all">
      <!-- Allow form submission with keyboard without duplicating the dialog button -->
      <input type="submit" tabindex="-1" style="position:absolute; top:-1000px">
      <div style="color: #ff0000 !important;" id="jqErrMsg" class="jqErrMsg"><br/>Test Message Test Message Test Message Test Message</div>
    </fieldset>
  </form>
</div>

</body>
</html>
