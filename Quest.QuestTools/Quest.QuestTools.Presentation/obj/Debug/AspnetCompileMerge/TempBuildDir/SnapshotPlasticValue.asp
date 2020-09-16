<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!--Date: February 4, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add Milvan warehouse
			 Updated to include option to download to excel
-->
<%
	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=PlasticInventoryReport.xls"
	Else
%>
<style>
	body { font-family: arial; }
	td { font-size: 13px; }
</style>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%
End If
ReportName = Request.Querystring("ReportName")
Country = Request.Querystring("Country")
ReportMonth = Left(Right(ReportName,6),2)
IF Left(ReportMonth,1) = 0 Then
	ReportMonth = Right(ReportMonth,1)
end if
ReportYear = Right(ReportName,4)
RecordNumbers = 0

' Canada View NASHUA USA view JUPITER - for future reports
if Country ="USA" then
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'JUPITER') order by PART ASC"
else
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP' OR WAREHOUSE = 'MILVAN') order by PART ASC"
end if


Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection2

'Create a Query
SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
Set RS2 = GetDisconnectedRS(SQL2, DBConnection)
'Get a Record Set

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<h2> Plastic Inventory As of  <%Response.write ReportMonth%> / <%Response.write ReportYear%></h2>
<h3> Using Inventory from: <%Response.write ReportName%></h3>
<h3> Viewing <%Response.write Country%></h3>

<% If Request("Download") <> "YES" Then %>
<a href="SnapshotPlasticValue.asp?Country=<%response.write Country%>&ReportName=<%response.write ReportName%>&Download=YES" target="_self"><b>Download Excel Copy</a><br/>
<% End If %>

<table border='1' class='sortable'>
  <tr><th>Part</th>
	<th>Description</th>
    <th>Qty</th>
    <th>Length (ft)</th>
    <th>Colour</th>
	<th>Bundle</th>
	<th>Entry Date</th>
    <th>LBF</th>
	<% 
	if Country ="USA" then
	%>
		<th>ExchangeRate <br> (CAD per USD)</th>
	<%
	end if
	%>

	<th>Value
	<% 
	if Country ="USA" then
		Response.write " ($USD)"
	else  
		Response.write " ($CAD)"
	end if
	%>
	</th>
	</tr>

<%

rs.movefirst
do while not rs.eof

invpart = rs("part")
rs2.filter =""
rs2.filter = "Part ='" & invpart & "'"
errormsg = 1
NotPl = 1
if rs2.eof then
		description = "N/A - Not Found in Master"
else
		description = rs2("Description")
		lbf = rs2("lbf")
		errormsg = 0
		if rs2("inventoryType") = "Plastic" then
			NotPl = 0
		else 
			NotPl = 1
		end if	
end if

partqty = rs("Qty")
lmm = CDBL(rs("lmm"))
lft = CDBL(rs("lft"))
linch = CDBL(rs("linch"))
colour = rs("colour")
project = rs("project")
bundle = rs("bundle")

if errormsg = 1 and NotPl =0 then
response.write invpart & " not in inventory master - " & rs("id") & " <BR>"
end if
if NotPl =  0 then
'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

EnterDate = rs("DateIn")


if Country ="USA" then
	PartMonth = Month(EnterDate)
	if Len(PartMonth) = 1 then
		PartMonth = "0" & PartMonth
	end if
	PartYear = Year(EnterDate)
	
	str_Period = PartYear & PartMonth
	if Len(str_Period) = 6 then
	else 
			'Use Default Exchange Rate Set by Mahesh April 2019 - 1.3
			ExchangeRate = 1.3
	end if
	
	Dim cn_SQL: Set cn_SQL = Server.CreateObject("adodb.connection")
	
	DBOpen cn_SQL, True
	
	SQL2 = "SELECT * FROM _qws_Inv_SupplierPrices order BY Period DESC"
	Set rs_Prices = GetDisconnectedRS(SQL2, cn_SQL)
	
	rs_Prices.Filter = "Period=" & str_Period + 0
	If Not rs_Prices.EOF Then
			ExchangeRate = rs_Prices.Fields("ExchangeRate")
	ELSE
			'Use Default Exchange Rate Set by Mahesh April 2019 - 1.3
			ExchangeRate = 1.3
	End if

	
else
	ExchangeRate = 1
end if

if ExchangeRate = 0 then
	ExchangeRate = 1
end if

	if lbf > 5 then 
	pricebar = partqty * lbf
	value2 = value2 + pricebar
	else
			tempvalue = (partqty * (lbf * lft ))
	value = value + tempvalue
	end if

'	if colour = "White" Then
'		paintlft = (linch/12) * partqty
'		totalpaintlft = totalpaintlft + paintlft
'	else
'	
'		If colour = "Mill" then
'		else
'			paintlft2 = (linch/12) * partqty
'			totalpaintlft2 = totalpaintlft2 + paintlft2
'		end if
'	
'	end if

%>

  <tr>
	<td><% response.write invpart %></td>
	<td><% response.write description %></td>
	<td><% response.write partqty %></td>
    <td><% response.write lft %></td>
    <td><% response.write colour %></td>
	<td><% response.write bundle %></td>
	<td><% response.write EnterDate %></td>
    <td><% response.write lbf%></td> 
	<% 
		if Country ="USA" then
		%><td><% response.write ExchangeRate %></td><%
		end if
	%>
    <td>$<% if lbf > 5 then
	response.write round(pricebar / ExchangeRate,2)
	else
	response.write round(tempvalue / ExchangeRate ,2)
	end if %></td>
    </tr>

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

	end if ' End NotPl code

rs.movenext
loop

%> </Table> <%
'Plastics do not have a colour - Shaun December 2014
'paintvalue1 = (totalpaintlft *0.14)
'paintvalue2 = (totalpaintlft2 *0.38)

response.write "<BR><HR>"
response.write "Total lbf Material $" & round(value / ExchangeRate,2) & "<BR>"
response.write "Total perbar Material $" & round(value2 /ExchangeRate ,2)  & "<BR>"
response.write "<HR>"
'response.write "SubTotal = $" & round(value,2) + round(value2,2) & "<BR><BR>"
'response.write "Total Paint White $" & paintv & "<BR>"
'response.write "Total Paint Project $" & paintvalue2 & "<BR>"
'response.write "<HR>"
response.write "Grand Total = $" & round(((value + value2) /ExchangeRate),2) 
'+ round(paintvalue1,2) + round(paintvalue2,2)

%>
<% 

rs.close
set rs=nothing
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>
</body>
</html>



