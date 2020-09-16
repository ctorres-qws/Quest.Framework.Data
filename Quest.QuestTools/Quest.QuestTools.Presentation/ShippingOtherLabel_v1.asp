<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<title>Non-Window Label</title>
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
job_name=UCASE(request.querystring("job_name"))
floor_name=UCASE(request.querystring("floor_name"))
tag_name=UCASE(request.querystring("tag_name"))
qty_name=UCASE(request.querystring("qty_name"))
desc_name=UCASE(request.querystring("desc_name"))
NewBarcode = ""

%>

<table align= "center" valign ="top" width="500" cellspacing="1" cellpadding="1" Style="font-size: 190%;">
	<tr>
		<td Style="font-size: 60%;" align = 'center' valign = 'top' colspan = "3" ><b><%response.write QTY_name & " " & Desc_Name %> </b></td>
	</tr>
	<tr>
		<td align = 'center' valign = 'center'>Job: <b><%response.write Job_name %></b></td>
		<td align = 'center' valign = 'center'>Floor: <b><%response.write Floor_Name %></b></td>
		<td align = 'center' valign = 'center'>Tag: <% response.write Tag_Name %></td>
	</tr>
	<tr> 
	<%
	NewBarcode = "00" & Job_name & Floor_Name & "-"& Tag_Name & ":"  & QTY_name & " " & Desc_Name
	%>
		<td Style="font-size: 60%;" align = 'center' valign = 'top' ><img align = 'center'  src="http://chart.apis.google.com/chart?cht=qr&chs=100x100&chl=<% response.write UCASE(NewBarcode) %>&chld=H|0" alt="Barcode" /></td>
		<td align = 'center' valign = 'center'> <b><%response.write Date %></b></td>
		<td Style="font-size: 60%;" align = 'center' valign = 'top' ><img align = 'center'  src="http://chart.apis.google.com/chart?cht=qr&chs=100x100&chl=<% response.write UCASE(NewBarcode) %>&chld=H|0" alt="Barcode" /></td>
	
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
