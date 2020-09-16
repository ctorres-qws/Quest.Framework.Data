<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
<title>Label Printer</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no" />
<style type="text/css">
@page {
margin-left:1mm;
margin-right:1mm;
margin-top:4mm;
margin-bottom:1mm;
font-family: Helvetica, Arial, sans-serif;
 size: portrait;
}

@media print {
    html, body {
        height: 99%;    
    }
}

@media print
{
table {page-break-after:always}
}

#upsidedown {
	font-size:50px;
    -webkit-transform: rotate(180deg);
    -moz-transform: rotate(180deg);
    -ms-transform: rotate(180deg);
    -o-transform: rotate(180deg);
}

</style>
</head>

<body>
<%
Dim item
	
	Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT top 5 * from Z_GlassDB order by ID Desc"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
do while not rs.eof
		

%>

<table align= "center" frame="box"  height = "725px" width = "384px" cellspacing="1" cellpadding="1">

	<tr>
		<td><img align = 'right' id = "upsidedown" src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("Barcode"))) %>&chld=H|0" alt="Barcode" /></td>
		<td><img align = 'center' id = "upsidedown" src="qlogoV.jpg" width="150" height="40" /></td>
		<td><img align = 'left' id = "upsidedown" src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("Barcode"))) %>&chld=H|0" alt="Barcode" /></td>
	</tr>
	<tr>
		<td colspan="3" id = "upsidedown" align = 'center' valign = 'center' Style="font-size: 200%;" > <b><%response.write RS("Dim X") %>'' x <%response.write RS("Dim Y") %>'' </b></td>
	</tr>
	<tr>
		<td colspan="3" id = "upsidedown" align = 'center' valign = 'center'  Style="font-size: 200%;"> <sup Style="font-size: 25%;">JOB</sup>&nbsp;<b><%response.write RS("JOB") %> &nbsp;&nbsp; <%response.write RS("FLoor") %></b>&nbsp;<sup Style="font-size: 25%">FLOOR</sup></td>
	</tr>
	<tr>
		<td colspan="3" id = "upsidedown" align = 'center' valign = 'center'  Style="font-size: 600%;"><b><%response.write rs("TAG") %></b><sup Style="font-size: 10%;">TAG</sup></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td colspan="3"   align = 'center' valign = 'center'  Style="font-size: 600%;"><b><%response.write rs("TAG") %></b><sup Style="font-size: 10%;">TAG</sup></td>
	</tr>
	<tr>
	<td colspan="3"   align = 'center' valign = 'center'  Style="font-size: 200%;"> <sup Style="font-size: 25%;">JOB</sup>&nbsp;<b><%response.write RS("JOB") %> &nbsp;&nbsp; <%response.write RS("FLoor") %></b>&nbsp;<sup Style="font-size: 25%">FLOOR</sup></td>
	</tr>
	<tr>
	<td colspan="3"   align = 'center' valign = 'center'  Style="font-size: 200%;"> <b><%response.write RS("Dim X") %>'' x <%response.write RS("Dim Y") %>'' </b></td>
	</tr>
	<tr>
		<td><img align = 'right'  src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("Barcode"))) %>&chld=H|0" alt="Barcode" /></td>
		<td><img align = 'center'  src="qlogoV.jpg" width="150" height="40" /></td>
		<td><img align = 'left'  src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("Barcode"))) %>&chld=H|0" alt="Barcode" /></td>
	</tr>


	
	
</table>

<%
rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>
<script>
//window.print()
</script>
</body>
</html>