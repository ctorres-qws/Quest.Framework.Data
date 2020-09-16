<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<title>Sheet Bundle Label</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">

body { margin-top: 5px; }

body,td,th {
	font-family: Verdana, Geneva, sans-serif;
	font-weight: normal;
	font-size: 18px;
}

<!--@media print
{
table {page-break-after:always}
}
-->
</style>

</head>

<body  align= "center" valign ="top" link="#000000" vlink="#C0C0C0" alink="#F6F000">

<%
'Value collected for printing in SHippingOtherLabelEnter.asp
colorcode =UCASE(request.querystring("colorcode"))
jobCode=UCASE(request.querystring("JobCode"))
SID=UCASE(request.querystring("SID"))

workOrder=UCASE(request.querystring("WorkOrder"))
PurchaseOrder=UCASE(request.querystring("PurchaseOrder"))
Datein=UCASE(request.querystring("Datein"))

PartNumber=UCASE(request.querystring("PartNumber"))
Size=UCASE(request.querystring("Size"))
Qty=request.querystring("Qty")

LabelPrintDate = Now



%>

<table align= "center" valign ="top" width="1000" cellspacing="1" cellpadding="1" border ="1" Style="font-size: 190%;">

	<tr>
		<td align = 'center' valign = 'center' width="400">Color: <b><%response.write ColorCode %></b></td>
		<td align = 'center' valign = 'center' width="300">Job: <b><%response.write jobCode %></b></td>
		<td align = 'center' valign = 'center' width="300">ID: <b><% response.write SID %></b></td>
	</tr>
	<tr>
		<td align = 'center' valign = 'center'>Work Order: <b><%response.write WorkOrder %></b></td>
		<td align = 'center' valign = 'center'>PO: <b><%response.write PurchaseOrder %></b></td>
		<td align = 'center' valign = 'center'>Date: <b><% response.write Datein %></b></td>
	</tr>
	<tr>
		<td align = 'center' valign = 'center'>Part #: <b><%response.write PartNumber%></b></td>
		<td align = 'center' valign = 'center'>Size: <b><%response.write Size %></b></td>
		<td align = 'center' valign = 'center'>Qty: <b><% response.write QTY %></b></td>
	</tr>
	<tr>
		<td align = 'center' valign = 'center'>Printed: <b><%response.write LabelPrintDate %></b></td>
		<td align = 'center' valign = 'center'></td>
		<td align = 'center' valign = 'center'><b>Quest Window Systems</b></td>
	</tr>
	</tr>	
	<tr>
		<td Style="font-size: 60%;" align = 'center' valign = 'top' colspan = "3" ><b><%response.write NewBarcode %> </b></td>
	</tr>

	</table>
	<div style="page-break-after: always;"></div>
	<div>&nbsp;
	</div>

</body>
</html>
