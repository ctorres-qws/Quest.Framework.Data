<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Date: January 13, 2019
	Modified By: Michelle Dungo
	Changes: Modified to generate cycle count viewer for hardware
-->
<style>
	body { font-family: arial; }
	td { font-size: 13px; }
</style>
<%
StartDate = Request.Querystring("startDate") & ""
EndDate = Request.Querystring("EndDate") & ""

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_HARDWARE WHERE TRIM(UCASE([MovedNote])) = 'CC' AND PART IN (SELECT PART FROM Y_HARDWARE_MASTER) AND LASTMODIFY >= #"&StartDate&"# AND LASTMODIFY <= #"&EndDate&"# ORDER BY AISLE ASC, RACK ASC"

rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
<%
	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=HardwareInventoryAdjustmentReport.xls"
	End If
%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
<script >
 table {
border-collapse:collapse;
}
</script>
 <script src="sorttable.js"></script>
<h3> Modified Date Range From: <%Response.write StartDate%> To <%Response.write EndDate%></h3>	
<body>
<% If Request("Download") <> "YES" Then %>
<a href="CycleCount_InventoryViewerHardware.asp?startDate=<%response.write StartDate%>&endDate=<%response.write EndDate%>&Download=YES" target="_self"><b>Download Excel Copy</a><br/>
<% End If %>
<%
response.write "Click on the Headers of each column to sort Ascending/Descending  "
response.write "<table border='1' class='sortable'><tr><th>Aisle</th><th>Rack</th><th>Level</th><th>Part</th><th>Quantity</th><th>PO</th><th>Enter Date</th><th>Modify Date</th><th>Warehouse</th><th>Moved Note</th></tr>"
do while not rs.eof
		response.write "<tr><td>" & RS("Aisle") & "</td><td>" & RS("Rack") & "</td><td>" & RS("Level") &"</td><td>" & RS("Part") & "</td><td>" & RS("QTY") & "</td><td>" & RS("PO") & "</td>" 
		response.write "<td>" & RS("EnterDate") & "</td><td>" & RS("LastModify") & "</td><td>" & RS("Warehouse") & "</td><td>" & RS("MovedNote") & "</td>"
		response.write " </tr>"
	rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
</table> 
</body>

