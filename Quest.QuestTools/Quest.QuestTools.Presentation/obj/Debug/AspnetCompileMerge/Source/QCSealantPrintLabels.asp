<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 25th, by Michael Bernholtz - Label to be Printed for Sealants-->
<!-- QC_SEALANT Table created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- One of 3 Tables - QC_GLASS, QC_SPACER, QC_SEALANT-->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
<title>Label Printer</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no" />

</head>


<%

Dim InventoryType, ItemID
'InventoryType = Request.Querystring("InventoryType")
QCID = 0
QCID = Request.QueryString("QCid")

Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT MSE.ItemName, SE.Identifier, SE.EntryDate, SE.printed, SE.Id FROM QC_SEALANT AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID ORDER BY MASTERID ASC"
		rs.Cursortype = GetDBCursorType
		rs.Locktype = GetDBLockType
		rs.Open strSQL, DBConnection
		rs.filter = "ID= " & QCID 

if QCID <> 0 then
		
	ItemName = rs("ItemName")
	Identifier = rs("Identifier")
	EntryDate = rs("EntryDate")

	rs("Printed")= 1
	rs.update
end if
%>

<body>


<br>
<table align= "center" frame="box" width="80pt" cellspacing="1" cellpadding="1">
	
	<tr>
		<td colspan="2" align = 'center' style="font-size: 125%;"> <b><%response.write ItemName %></b></td>
	</tr>
	<tr>
		
		<td align = 'center' style="font-size: 125%;"> <b><%response.write Identifier %></b></td>
		<td align = 'center' style="font-size: 125%;"> <b><%response.write EntryDate %><b></td>
	</tr>
	<tr>
		<td  align = 'center'><img src="qlogoV.jpg" width="150" height="40" /></td>
		<!--	<td align = 'center'><img src="http://www.scandit.com/barcode-generator/?symbology=code39&value=<% response.write UCASE(trim(rs("Identifier"))) %>&ec=L&size=100" alt="Barcode" /></td>-->
		<td align = 'center'><img src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("Identifier"))) %>&chld=H|0" alt="Barcode" /></td>
		
		
	</tr>
</table>
<script>
window.print()
</script>
<p>
 
</p>
<p>&nbsp;</p>
<%
rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>

</body>
</html>