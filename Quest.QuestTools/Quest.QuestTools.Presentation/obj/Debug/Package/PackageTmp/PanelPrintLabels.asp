<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created October 2018 Michael Bernholtz - Label to be Printed for Sheets received from Pending-->
<!-- QC_Glass Table created for David Ofir and Jody Cash, Implemented by Michael Bernholtz-->
<!-- Z_Jobs LabelPrint Column Set to No when Pending sent to Nashua, Set to Yes when printed-->


<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
<title>Label Printer</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no" />

</head>


<% 

PID = 0
PID = Request.QueryString("Pid")

Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Y_INV WHERE ID = " & PID & " ORDER BY ID ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
if PID <> 0 then
	
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM Y_COLOR WHERE PROJECT = '" & RS("colour") & "' ORDER BY ID ASC"
	rs2.Cursortype = GetDBCursorType
	rs2.Locktype = GetDBLockType
	rs2.Open strSQL2, DBConnection
	ColourCode = rs2("Code")
	rs2.close
	set rs2 = nothing
	
	PartName = rs("part")
	SheetSize = rs("Width") & " X " & rs("Height") 
	ColorPO = rs("ColorPO")
	Qty = rs("qty")
	
	rs("LabelPrint")= "Yes"
	rs.update
end if
%>

<body>

<%
Count = 1
Do until Count >QTY
%>

<table align= "center" frame="box" width="300px" cellspacing="1" cellpadding="1">
	
	<tr>
		<td align = 'center' style="font-size: 125%;"> <b><%response.write PartName %></b></td>
		<td align = 'center' style="font-size: 125%;"> <b><%response.write SheetSize %><b></td>
	</tr>
	
	<tr>
		<td align = 'center' style="font-size: 125%;"> <b><%response.write ColourCode %></b></td>
		<td align = 'center' style="font-size: 75%;"> <b>Color PO: <%response.write ColorPO %><b></td>
	</tr>
	<tr>
		<td colspan="2" align = 'center'><img src="http://chart.apis.google.com/chart?cht=qr&chs=75x75&chl=<% response.write UCASE(trim(rs("ID"))) %>&chld=H|0" alt="Barcode" /></td>

	</tr>
	
</table>
<br>

<%
	if Count = Qty then
	else
%>
<div style="page-break-after: always;"></div>
<%
	end if
%>
<%
Count = Count + 1
Loop
%>


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