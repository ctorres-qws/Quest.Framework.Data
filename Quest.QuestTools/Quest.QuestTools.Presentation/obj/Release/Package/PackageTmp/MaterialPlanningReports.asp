<!--#include file="dbpath.asp"-->
<%
	Dim str_MsgErr
	Dim cn_SQL, rs_Data
	Dim str_Job, str_Floors, str_Windows

	str_Job = Request("Jobs")
	str_Report = Request("Reports")
	str_Colour = Request("Colours")
	str_PageAction = Request("PageAction")

	If UCase(str_PageAction) <> "SELECT_COLOUR" Then
		str_Colour = ""
	End If

	Set cn_SQL = Server.CreateObject("ADODB.Connection")
	cn_SQL.Open GetConnectionStr(true)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Quest Dashboard</title>
	<meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
	<link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Request from Jody Cash on January 9th 2014 to change the Auto Refresh from 1200 to 90 -->
	<link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	<link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css?v=1"  type="text/css"/>
	<script type="application/x-javascript" src="/iui/iui.js"></script>
<link rel="stylesheet" href="javascript/jquery-ui.css">
<script type="text/javascript" src="javascript/jquery-1.11.3.js"></script>
<script type="text/javascript" src="javascript/jquery-ui.js"></script>
	<script type="text/javascript">
		iui.animOn = true;
	</script>
	<style>
		.csForm { background-color: #eaeaea; border: 2px solid #cccccc; border-radius: 5px; width: 600px !important; position: fixed; top: 100px !important; z-index: 9999;}
		.csFormContainer { padding: 50px; }

		.csButton { height: 30px !important; padding: 10px 8px;}

		.csLabel { width: 120px; text-align: right; }

		select { font-size: 22px; }

		input[type='text'], select  {
			margin: 0px 5px 5px 5px;
			padding: 0px 0px 0px 0px !important;
			border-radius: 4px;
			border: 1px solid rgb(200, 200, 200);
			border-image: none;
			text-align: left !important;
			height: 40px;
			width: 300px !important;
		}
	
		table, tr { border-bottom: 0px !important; }

.ui-state-hover, .ui-autocomplete li:hover
{
    background: rgb(238,238,238);
    margin-left: 1px;
    margin-right: 1px;
}

.ui-autocomplete {
  font-weight: normal;
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 1000;
  float: left;
  display: none;
  min-width: 280px;
  width: 500px !important;
  padding: 4px 0;
  margin: 2px 0 0 0;
  list-style: none;
  background-color: #ffffff;
  border-color: #ccc;
  border-color: rgba(0, 0, 0, 0.2);
  border-style: solid;
  border-width: 1px;
  -webkit-border-radius: 5px;
  -moz-border-radius: 5px;
  border-radius: 5px;
  -webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  -webkit-background-clip: padding-box;
  -moz-background-clip: padding;
  background-clip: padding-box;
  *border-right-width: 2px;
  *border-bottom-width: 2px;
}

.csTable { border: 1px solid #999999; font-size: 13px; }

.csTable tr { height: 28px; }

.csTable tr:nth-child(odd){
  background-color: #eaeaea;
  color: #0;
}

.csTableHdr {background-color: #cccccc !important;}
.csTableHdrCalc {background-color: #eeeeee !important;}
.csTable td {text-align: right; padding-left: 10px; padding-right: 10px;}

.csSep {border-top: 1px solid #cccccc;}

	.csOther { color: #999999; }

	</style>
</head>
<body>
	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Home</a>
	</div>

<ul id="screen1" title="Material Planning" selected="true">
	<form method="post" name="fMain">
	<input type="hidden" name="AutoPost" value="TRUE">
	<input type="hidden" name="PageAction" value="">
	<input type="hidden" name="MaterialID" value="">

<li>
	<div style="padding-left: 20px;">
	<table cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td style="width: 160px; ">Select a Report:&nbsp;</td>
			<td style="padding-left: 20px;">
				<select name="Reports" class="jqReports">
					<option value="JOB" <% If str_Report = "JOB" Then Response.Write(" selected ") %>>Job Report</option>
					<option value="PART" <% If str_Report = "PART" Then Response.Write(" selected ") %>>Parts Report</option>
				</select>
			</td>
			<td>&nbsp;<div style='display: none;'><a class="csButton" type="cancel" href="javascript: void()" onclick="RefreshPage();" target="_self" title="Refresh Page">Refresh</a></div></td>
		</tr>
	</table>
	</div>
</li>
<%
	If str_Report = "" Or str_Report = "JOB" Then
%>
	<li>
		<div style="padding-left: 20px;">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td style="width: 160px;">Select a Job:</td>
				<td style="padding-left: 20px;">
					<select name="Jobs" class="jqJobs">
						<option>Select Job</option>
<%

Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT * FROM _qws_Planning_Jobs ORDER BY JobName ASC", cn_SQL

Do While Not rs_Data.EOF
	Dim str_Selected: str_Selected = ""
	If str_Job = rs_Data.Fields("JobName").Value Then
		str_Selected = " selected='selected' "
		str_Floors = rs_Data.Fields("Floors").Value
		str_Windows = rs_Data.Fields("Windows").Value
	End If
	Response.Write("<option value='" & rs_Data.Fields("JobName").Value & "' " & str_Selected & ">" & rs_Data.Fields("JobName").Value & "</option>")
	rs_Data.MoveNext
Loop

rs_Data.Close()
Set rs_Data = Nothing

%>
				</select>
				</td>
				<td>
			<select name="Colours" class="jqColours">
<%
If str_Job <> "" Then

Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT * FROM Y_Color WHERE Job='" & str_Job & "' ORDER BY Project ASC", cn_SQL

If str_Colour = "" Then
	If Not rs_Data.EOF Then
		str_Colour = rs_Data("Project")
	End If
End If

Do While Not rs_Data.EOF
	str_Selected = ""
	If UCase(rs_Data("Project")) = UCase(str_Colour) Then str_Selected = " selected "
	Response.Write("<option value='" & rs_Data("Project") & "' " & str_Selected & ">" & rs_Data("Project") & "</option>")
	rs_Data.MoveNext
Loop
rs_Data.Close()
Set rs_Data = Nothing
End If
%>
			</select>
				</td>
			</tr>
		</table>
	</div>
	</li>

<li>
<%
	Set rs_Data = Server.CreateObject("ADODB.Recordset")
	rs_Data.CursorType = GetDBCursorType
	rs_Data.LockType = GetDBLockType
	rs_Data.Open "SELECT * FROM _qws_Planning_Jobs WHERE JobName='" & str_Job & "'", cn_SQL

	If Not rs_Data.EOF Then
		str_Floors = rs_Data("Floors")
		str_Windows = rs_Data("Windows")
	End If
	rs_Data.Close: Set rs_Data = Nothing

%>
<div style="padding-left: 20px;">
Project: <%= str_Job %><br/>
Floors: <%= str_Floors %><br/>
Windows: <%= str_Windows %><br/>
Date: <%= Now %><br/>
</div>
</li>

<li>
<div style="padding-left: 20px;">
<table cellspacing="0" cellpadding="0" class="csTable">
	<tr class="csTableHdr">
		<td>&nbsp;Part</td>
		<td style="text-align: left;">Description</td>
		<td class="csTableHdrCalc"><%= str_Job %> Qty<br/>Estimate</td>
		<td class="csTableHdrCalc">-</td>
		<td class="csTableHdrCalc"><%= str_Job %> Qty<br/>Consumed</td>
		<td class="csTableHdrCalc">=</td>
		<td class="csTableHdrCalc"><%= str_Job %> Qty<br/>Remaining</td>
		<td>Qty<br />Nashua</td>
		<td class="csTableHdrCalc"><%= str_Job %> Qty<br/>Short</td>
		<td tip="Same Colour code used on other projects.">Qty Other<br />Nashua</td>
		<td>Other Projects<br/>Qty Estimate</td>
		<td>Other Projects<br/>Qty Consumed</td>
		<td>Other Projects<br/>Qty Remaining</td>
		<td>Qty<br />Nashua</td>
		<td>Qty<br />Goreway</td>
		<td>Qty<br />Horner</td>
		<td>Qty<br />Metra</td>
		<td>Qty<br />Durapaint</td>
		<td>Qty<br />DurapaintWIP</td>
		<td>Qty&nbsp;<br />Sapa&nbsp;</td>
		<td>Total Qty<br />Warehouses</td>
	</tr>
<%
	Dim str_SQL

	If str_Colour <> "" Then
	str_SQL = ReadFile(Server.MapPath("MaterialsPlanningReport_Job.txt"))
	str_Colour2 = SplitArgs(str_Colour ," ", 1)
	str_SQL = Replace(str_SQL, "{0}", str_Colour)
	str_SQL = Replace(str_SQL, "{1}", str_Colour2)
	Set rs_Data = Server.CreateObject("ADODB.Recordset")
	rs_Data.CursorType = GetDBCursorType
	rs_Data.LockType = GetDBLockType
	rs_Data.Open str_SQL, cn_SQL

	Do While Not rs_Data.EOF
		i_InvQty = rs_Data("Qty_Nashua")
		i_Qty = rs_Data("EstQtyProject") - rs_Data("ConsumedQty")
		If i_Qty < 0 Then
			i_Qty = rs_Data("EstQtyProject")
		End If
		i_Diff = i_InvQty - i_Qty
		str_Diff = ""
		If i_Diff < 100 Then
			str_Diff = FormatNumber(Abs(i_Diff),0)
		End If
		i_OtherConsumedQty = rs_Data("OtherConsumedQty")
		If rs_Data("EstQtyOther") = 0 Then i_OtherConsumedQty = 0
%>
	<tr>
		<td><nobr><%= rs_Data("Part") %></nobr></td>
		<td style="text-align: left; font-size: 11px;"><%= rs_Data("Description") %></td>
		<td style="width: 5%; border-left: 1px solid #cccccc;"><%= FormatNumber(rs_Data("EstQtyProject"),0) %></td>
		<td></td>
		<td style="width: 5%;"><%= FormatNumber(rs_Data("ConsumedQty"),0) %></td>
		<td></td>
		<td style="width: 5%; font-weight: bold; border-right: 1px solid #cccccc;"><%= FormatNumber(rs_Data("EstQtyProject") - rs_Data("ConsumedQty"),0) %></td>
		<td><%= FormatNumber(rs_Data("Qty_Nashua"),0) %></td>
		<td style="width: 5%;border-left: 1px solid #cccccc; bold; border-right: 1px solid #cccccc; <% If str_Diff <> "" Then Response.Write(" background-color: rgb(255,220,142); ") %>"><%= str_Diff %></td>
		<td class=""><%= FormatNumber(rs_Data("Qty_Nashua_Other"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("EstQtyOther"),0) %></td>
		<td class="csOther"><%= FormatNumber(i_OtherConsumedQty,0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("EstQtyOther") - i_OtherConsumedQty,0) %></td>
		
		<td style="width: 5%; border-left: 1px solid #cccccc;" class="csOther"><%= FormatNumber(rs_Data("Qty_Nashua"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Goreway"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Horner"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Metra"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Durapaint"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_DurapaintWIP"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Sapa"),0) %>&nbsp;</td>
		<td class="csOther"><%= FormatNumber(rs_Data("TotalWarehouseQty"),0) %></td>
	</tr>
<%
		rs_Data.MoveNext
	Loop
	rs_Data.Close()
	Set rs_Data = Nothing
End If
%>
</table>
<div>
</li>
<%
End If
' -------------------------------------------------------------------------------------------- PART
If str_Report = "PART" Then
%>
	<li>
		<div style="padding-left: 20px;">
		<table cellpadding="0" cellspacing="0" border="0" >
			<tr>

				<td style="width: 160px;">Colour:&nbsp;</td>
				<td style="padding-left: 20px;">
					<select name="PartColour" class="jqPartColours">
						<option value="Ext." <% If Request("PartColour") = "Int." Then Response.Write(" selected='selected' ") %>>Exterior</option>
						<option value="Int." <% If Request("PartColour") = "Int." Then Response.Write(" selected='selected' ") %>>Interior</option>
					</select>
				</td>
				<td style="width: 160px;">Difference:</td>
				<td style="display: none;"><input type="checkbox" value="SHORT" name="FilterShort">&nbsp;Only show short</td>
				<td style="padding-left: 20px;">
					<select name="Difference" style="width: 100px !important;">
						<option value=""></option>
						<option value="0" <% If Request("Difference") = "0" Then Response.Write(" selected='selected'") %>>< 0</option>
						<option value="50" <% If Request("Difference") = "50" Then Response.Write(" selected='selected'") %>>< 50</option>
						<option value="100" <% If Request("Difference") = "100" Then Response.Write(" selected='selected'") %>>< 100</option>
						<option value="200" <% If Request("Difference") = "200" Then Response.Write(" selected='selected'") %>>< 200</option>
						<option value="300" <% If Request("Difference") = "300" Then Response.Write(" selected='selected'") %>>< 300</option>
						<option value="400" <% If Request("Difference") = "400" Then Response.Write(" selected='selected'") %>>< 400</option>
						<option value="500" <% If Request("Difference") = "500" Then Response.Write(" selected='selected'") %>>< 500</option>
						<option value="1000" <% If Request("Difference") = "1000" Then Response.Write(" selected='selected'") %>>< 1000</option>
					</select>
				</td>
				<td><div style='width: 50px;'></div></td>
				<td><a class="csButton" type="cancel" href="javascript: void()" onclick="RefreshPage();" target="_self" title="Filter">Filter</a></td>
			</tr>
		</table>
		</div>
	</li>
<li>
<div style="padding-left: 20px;">
<table cellspacing="0" cellpadding="0" style="width: 99%;" class="csTable">
	<tr class="csTableHdr">
		<td>&nbsp;Part</td>
		<td>Projects<br/>Qty Estimate</td>
		<td>Projects<br/>Qty Consumed</td>
		<td>Projects<br/>Qty Remaining</td>
		<td style='text-align: center;'>Difference</td>
		<td>Qty<br />Nashua</td>
		<td>Total Qty<br />Warehouses</td>
		<td>Qty<br />Goreway</td>
		<td>Qty<br />Horner</td>
		<td>Qty<br />Metra</td>
		<td>Qty<br />Durapaint</td>
		<td>Qty<br />DurapaintWIP</td>
		<td>Qty&nbsp;<br />Sapa&nbsp;</td>
		
	</tr>
<%


	str_SQL = ReadFile(Server.MapPath("MaterialsPlanningReport_Part.txt"))
	str_SQL = Replace(str_SQL, "{0}", Request("PartColour"))
	Set rs_Data = Server.CreateObject("ADODB.Recordset")
	rs_Data.CursorType = GetDBCursorType
	rs_Data.LockType = GetDBLockType
	rs_Data.Open str_SQL, cn_SQL

	Do While Not rs_Data.EOF
		i_Diff = 0: i_Qty = 0
		b_Display = True
		str_Color = ""

		i_InvQty = rs_Data("Qty_Nashua")
		i_Qty = rs_Data("ProjectsQty") - rs_Data("ConsumedQty")
		If i_Qty < 0 Then
			i_Qty = rs_Data("ProjectsQty")
		End If
		i_Diff = i_InvQty - i_Qty

		'str_Color = " style='background-color: #FFCC66 ;'"

		str_Diff = ""
		'If i_Diff < 100 Then
			str_Diff = FormatNumber(i_Diff,0)
		'End If

		If Request("Difference") <> "" Then
			b_Display = False
			If i_Diff < CLng(Request("Difference")) Then
				b_Display = True
			End If
		Else
			If i_Diff > 1000 Then
				str_Diff = ""
			End if
		End If

		If b_Display Then
%>
	<tr <%= str_Color %>>
		<td>&nbsp;<%= rs_Data("Part") %></td>
		<td><%= FormatNumber(rs_Data("ProjectsQty"),0) %></td>
		<td><%= FormatNumber(rs_Data("ConsumedQty"),0) %></td>
		<td><%= FormatNumber(rs_Data("ProjectsQty") - rs_Data("ConsumedQty"),0) %></td>
		<td <% If str_Diff <> "" Then Response.Write(" style='font-weight: bold; text-align: center;'") %>><%= str_Diff %></td>
		<td><%= FormatNumber(rs_Data("Qty_Nashua"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("TotalWarehouseQty"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Goreway"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Horner"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Metra"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Durapaint"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_DurapaintWIP"),0) %></td>
		<td class="csOther"><%= FormatNumber(rs_Data("Qty_Sapa"),0) %>&nbsp;</td>
	</tr>
<%
		End If

		rs_Data.MoveNext
	Loop
	rs_Data.Close()
	Set rs_Data = Nothing

%>
</table>
<div>
</li>

<%
End If
%>
<li>&nbsp;</li>
	</form>
</ul>
<br /><br />
	<script>

	$(document).ready(function() {

		$(".jqJobs").change(function() {
			fMain.PageAction.value = "SELECT_JOB";
			RefreshPage();
		});

		$(".jqReports").change(function() {
			fMain.PageAction.value = "SELECT_REPORT";
			RefreshPage();
		});

		$(".jqColours").change(function() {
			fMain.PageAction.value = "SELECT_COLOUR";
			RefreshPage();
		});

	});

	function RefreshPage() {
		fMain.submit();
	}

	</script>

<%

	Function SplitArgs(str_Args, str_Split, i_Index)
		Dim str_Ret
		a_Ret = Split(str_Args, str_Split)
		str_Ret = a_Ret(i_Index)
		SplitArgs = str_Ret
	End Function

%>
<%
DBConnection.close
Set DBConnection = nothing
%>
</body>
</html>
