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
margin-top:10mm;
margin-bottom:1mm;
font-family: Helvetica, Arial, sans-serif;
 size: landscape;
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

#vertical {
	font-size:50px;
    -webkit-transform: rotate(-90deg);
    -moz-transform: rotate(-90deg);
    -ms-transform: rotate(-90deg);
    -o-transform: rotate(-90deg);
}

#vertical2 {
	font-size:50px;
    -webkit-transform: rotate(90deg);
    -moz-transform: rotate(90deg);
    -ms-transform: rotate(90deg);
    -o-transform: rotate(90deg);
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

<table align= "center" frame="box"  width = "768px" height = "384px" cellspacing="1" cellpadding="1">

	<tr>
		<td rowspan="4" align = 'center' id='vertical2' width = 25% Style="font-size: 600%;"> <b><%response.write rs("TAG") %></b></td>
		<td align = 'left' > <b>JOB</b></td>
		<td align = 'right' > <b>FLOOR</b></td>
		<td rowspan="4" align = 'center' id='vertical'width = 25% Style="font-size: 600%;"> <b><%response.write rs("TAG") %></b></td>
	</tr>	
	
	<tr>
		<td align = 'left' Style="font-size: 300%;"> <b><%response.write RS("JOB") %> </b></td>
		<td  align = 'right' style="font-size: 300%;"> <b><%response.write RS("FLOOR") %> </b></td>
	</tr>
	<tr>
		<td align = 'left' width = 25%><img src="qlogoV.jpg" width="150" height="40" /></td>
		<td align = 'right' width = 25%><img src="http://chart.apis.google.com/chart?cht=qr&chs=100x100&chl=<% response.write UCASE(trim(rs("Barcode"))) %>&chld=H|0" alt="Barcode" /></td>
	</tr>

	<tr>
		<td colspan="2" align = 'center' style="font-size: 200%;" > <b><%response.write RS("Dim X") %>'' x <%response.write RS("Dim Y") %>'' </b></td>
		
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