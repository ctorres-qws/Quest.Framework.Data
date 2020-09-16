<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
	Dim str_Page: str_Page = "InventoryReportPrices.asp"
	Dim str_SearchYear
	Dim str_ID

	str_SearchYear = Request("SearchYear")
	If str_SearchYear = "" Then str_SearchYear = Year(Now)

	Dim str_QueryPeriod, str_PeriodYear, str_PeriodMonth
	str_QueryPeriod = Request("PeriodYear") & Request("PeriodMonth")
	If str_QueryPeriod  <> "" Then
		str_PeriodYear = Left(str_QueryPeriod, 4)
		str_PeriodMonth = Right(str_QueryPeriod, 2)
		str_SearchYear = str_PeriodYear
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
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script>

	var dialog
	$( document ).ready(function() {

		$("#add-price").click(function() {

			frmSupplierPrice.qwsAction.value = 'PRICE_UPDATE';
			frmSupplierPrice.ID.value = '';
			$(".jqPeriodYear").val($(".jqSearchYear").val());
			frmSupplierPrice.PriceSapaMill_H.value = '';
			frmSupplierPrice.PriceSapaMill_S.value = '';
			frmSupplierPrice.PriceSapaMontreal_H.value = '';
			frmSupplierPrice.PriceSapaMontreal_S.value = '';
			frmSupplierPrice.PriceExtal_H.value = '';
			frmSupplierPrice.PriceExtal_S.value = '';
			frmSupplierPrice.PriceMetra_H.value = '';
			frmSupplierPrice.PriceMetra_S.value = '';
			frmSupplierPrice.PriceCanArt_H.value = '';
			frmSupplierPrice.PriceCanArt_S.value = '';
			frmSupplierPrice.PriceApel_H.value = '';
			frmSupplierPrice.PriceApel_S.value = '';
			frmSupplierPrice.PriceDefault_H.value = '';
			frmSupplierPrice.PriceDefault_S.value = '';

			$(".jqErrMsg").html('');

			dialog.dialog( "open" );
			$("#dialog-form").css("height", "600");
		});

		dialog = $( "#dialog-form" ).dialog({
			autoOpen: false,
			height: 400,
			width: 650,
			modal: true,
			position: { my: 'top', at: 'top+100' },
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
			frmSupplierPrice.submit();
			dialog.dialog( "close" );
		}

	});

	function showSearchYear() {
		frmSearch.SearchYear.value = prices.SearchYear.options[prices.SearchYear.options.selectedIndex].value
		frmSearch.submit();
	}

	function editPrice(str_ID, str_Year, str_Month, str_PriceSapaMill_H, str_PriceSapaMill_S, str_PriceSapaMontreal_H, str_PriceSapaMontreal_S, str_PriceExtal_H, str_PriceExtal_S, str_PriceMetra_H, str_PriceMetra_S, str_PriceCanArt_H, str_PriceCanArt_S, str_PriceApel_H, str_PriceApel_S, str_PriceDefault_H, str_PriceDefault_S) {
		$(".jqErrMsg").html('');

		frmSupplierPrice.qwsAction.value = 'PRICE_UPDATE';
		frmSupplierPrice.ID.value = str_ID;
		frmSupplierPrice.PriceSapaMill_H.value = str_PriceSapaMill_H;
		frmSupplierPrice.PriceSapaMill_S.value = str_PriceSapaMill_S;
		frmSupplierPrice.PriceSapaMontreal_H.value = str_PriceSapaMontreal_H;
		frmSupplierPrice.PriceSapaMontreal_S.value = str_PriceSapaMontreal_S;
		frmSupplierPrice.PriceExtal_H.value = str_PriceExtal_H;
		frmSupplierPrice.PriceExtal_S.value = str_PriceExtal_S;
		frmSupplierPrice.PriceMetra_H.value = str_PriceMetra_H;
		frmSupplierPrice.PriceMetra_S.value = str_PriceMetra_S;
		frmSupplierPrice.PriceCanArt_H.value = str_PriceCanArt_H;
		frmSupplierPrice.PriceCanArt_S.value = str_PriceCanArt_S;
		frmSupplierPrice.PriceApel_H.value = str_PriceApel_H;
		frmSupplierPrice.PriceApel_S.value = str_PriceApel_S;
		frmSupplierPrice.PriceDefault_H.value = str_PriceDefault_H;
		frmSupplierPrice.PriceDefault_S.value = str_PriceDefault_S;

		$(".jqPeriodYear").val(str_Year);
		$(".jqPeriodMonth").val(str_Month);

		dialog.dialog( "open" );
		$("#dialog-form").css("height", "600");
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

	.ui-dialog .ui-dialog-content { overflow: hidden !important; }

	fieldset { border: 1px solid rgb(221,221,221); border-radius: 5px; }

	.csDialogRow { xborder: 1px solid black; }
	.csDialogRow > label { width: 200px; float: left !important; xborder: 1px solid black; }
	.csDialogRow > input { width: 100px; float: left !important; xborder: 1px solid black; margin-left: 0px !important; margin-top: 0px !important; }

	select { border: 1px solid rgb(221,221,221) !important; border-radius: 5px; padding: 3px 3px 3px 3px; margin-bottom: 5px; }

	</style>
</head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<div style=""><a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportSelectCostBySupplier.asp" target="_self" style="left: 90px;">Inventory Report</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPrices.asp" target="_self" style="left: 216px;">Supplier Prices</a></div>
		<div style=""><a class="button leftButton" type="cancel" href="InventoryReportPaintPrices.asp" target="_self" style="left: 333px;">Paint Pricing</a></div>
		<div style="left: 400px;"><a class="button rightButton" id="add-price" type="cancel" href="#" target="_self" style=""> + Add Supplier Price</a> </div>
	</div>

	<form id="prices" title="Inventory Prices" class="panel" name="prices" action="<%= str_Page %>" method="GET" target="_self" selected="true" autocomplete="off">
		<input name="qwsAction" value="" type="hidden">
		<h2>Enter Aluminium Price and Select a Snapshot of Inventory</h2>
		<fieldset>

		<div class="row">

			<label>Inventory Period</label>
			<div style="float: right; padding-right: 80px; padding-top: 12px; ">
			<select id='SearchYear' name="SearchYear" class="jqSearchYear">
<%
	If str_SearchYear <> "" Then Response.Write("<option value='" & str_SearchYear & "'>" & str_SearchYear & "</option>")
	For i = Year(Now) to 2010 Step Size - 1
		Response.Write("<option value='" & i & "'>" & i & "</option>")
	Next
%>
			</select>
			</div>
			&nbsp;<div style=""><a class="button rightButton" href="javascript: void()" onclick="showSearchYear();">Search</a></div>
		</div>

<%
	Dim str_Year: str_Year = str_SearchYear
	If str_Year = "" Then str_Year = Year(Now)

%>
	</fieldset>
	<BR>
<div class="toolbar">

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
		str_ID = rs_Data.Fields("ID")
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
	<td style="text-align: center;"><a href="javascript: void();" onclick="editPrice('<%= str_ID %>','<%= Left(str_Period,4) %>','<%= Right(str_Period,2) %>','<%= str_PriceSapaMill_H %>','<%= str_PriceSapaMill_S %>','<%= str_PriceSapaMontreal_H %>','<%= str_PriceSapaMontreal_S %>','<%= str_PriceExtal_H %>','<%= str_PriceExtal_S %>','<%= str_PriceMetra_H %>','<%= str_PriceMetra_S %>','<%= str_PriceCanArt_H %>','<%= str_PriceCanArt_S %>','<%= str_PriceApel_H %>','<%= str_PriceApel_S %>','<%= str_Default_H %>','<%= str_Default_S %>');">Edit</a></td>
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



<form name="frmEditPrice" action="<%= str_Page %> method="GET">
	<input type="hidden" name="PeriodYear">
	<input type="hidden" name="PeriodMonth">
	<input type="hidden" name="qwsAction">
</form>

<form name="frmSearch" action="<%= str_Page %>" method="GET">
	<input type="hidden" name="SearchYear">
	<input type="hidden" name="qwsAction">
</form>

<div id="dialog-form" title="Supplier Prices">
  <p class="validateTips">All form fields are required.</p>

  <form name="frmSupplierPrice" action="<%= str_Page %>" method="GET" target="_self" >
    <fieldset>
      <div>&nbsp;</div>
      <input type="hidden" name="qwsAction" value="PRICE_UPDATE">
      <input type="hidden" name="ID" value="">

<div class="csDialogRow"><label>Inventory Period</label>
			<select id='PeriodYear' name="PeriodYear" class="jqPeriodYear">
<%
	If str_PeriodYear <> "" Then Response.Write("<option value='" & str_PeriodYear & "'>" & str_PeriodYear & "</option>")
	For i = Year(Now) to 2010 Step Size - 1
		Response.Write("<option value='" & i & "'>" & i & "</option>")
	Next
%>
			</select>
			<select id='PeriodMonth' name="PeriodMonth" class="jqPeriodMonth">
<%
	If str_PeriodMonth <> "" Then Response.Write("<option value='" & Right("0" & str_PeriodMonth, 2) & "'>" & Right("0" & str_PeriodMonth, 2) & "</option>")
	For i = 1 to 12
		Response.Write("<option value='" & Right("0" & i, 2) & "'>" & Right("0" & i, 2) & "</option>")
	Next
%>
			</select>
</div>

<div class="csDialogRow"><label>Sapa Mill - Hollow</label><input type="number" name='PriceSapaMill_H' id='PriceSapaMill' value =""></div>

<div class="csDialogRow"><label>Sapa Mill - Solid</label><input type="number" name='PriceSapaMill_S' id='PriceSapaMill' value ="" ></div>

<div class="csDialogRow"><label>Sapa Montreal - Hollow</label><input type="number" name='PriceSapaMontreal_H' id='PriceSapaMontreal' value ="" ></div>

<div class="csDialogRow"><label>Sapa Montreal - Solid</label><input type="number" name='PriceSapaMontreal_S' id='PriceSapaMontreal' value ="" ></div>

<div class="csDialogRow"><label>Extal - Hollow</label><input type="number" name='PriceExtal_H' id='PriceExtal' value ="" ></div>

<div class="csDialogRow"><label>Extal - Solid</label><input type="number" name='PriceExtal_S' id='PriceExtal' value ="" ></div>

<div class="csDialogRow"><label>Metra - Hollow</label><input type="number" name='PriceMetra_H' id='PriceMetra' value ="" ></div>

<div class="csDialogRow"><label>Metra - Solid</label><input type="number" name='PriceMetra_S' id='PriceMetra' value ="" ></div>

<div class="csDialogRow"><label>Can-Art - Hollow</label><input type="number" name='PriceCanArt_H' id='PriceCanArt' value ="" ></div>

<div class="csDialogRow"><label>Can-Art - Solid</label><input type="number" name='PriceCanArt_S' id='PriceCanArt' value ="" ></div>

<div class="csDialogRow"><label>Apel - Hollow</label><input type="number" name='PriceApel_H' id='PriceApel' value ="" ></div>

<div class="csDialogRow"><label>Apel - Solid</label><input type="number" name='PriceApel_S' id='PriceApel' value ="" ></div>

<div class="csDialogRow"><label>Default - Hollow</label><input type="number" name='PriceDefault_H' id='PriceDefault' value ="" ></div>

<div class="csDialogRow"><label>Default - Solid</label><input type="number" name='PriceDefault_S' id='PriceDefault' value ="" ></div>

      <input type="submit" tabindex="-1" style="position:absolute; top:-1000px">
      <div style="color: #ff0000 !important;" id="jqErrMsg" class="jqErrMsg"><br/>Test Message Test Message Test Message Test Message</div>
    </fieldset>
  </form>
</div>

</body>
</html>
