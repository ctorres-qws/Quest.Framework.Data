<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add download button to convert to excel format
			Change price logic, to get value from column PPU if not empty or zero else retrieve price from master
	Date: February 4, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add Milvan warehouse			
-->
<%
	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=HardwareInventoryReport.xls"
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
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('NASHUA','MILVAN') order by PART ASC"
end if


Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection2

'Aluminum price
alumprice = AlumPrice

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<h2> Hardware Inventory As of  <%Response.write ReportMonth%> / <%Response.write ReportYear%></h2>
<h3> Using Inventory from: <%Response.write ReportName%></h3>
<h3> Viewing <%Response.write Country%></h3>

<% If Request("Download") <> "YES" Then %>
<a href="SnapshotHardwareValue.asp?Country=<%response.write Country%>&ReportName=<%response.write ReportName%>&Download=YES" target="_self"><b>Download Excel Copy</a><br/>
<% End If %>

<table border='1' class='sortable'>
	<tr>
	<th>Part</th>
	<th>Description</th>
    <th>Qty</th>
    <th>PO</th>
	<th>Price ($CAD)</th>
	<% 
	if Country ="USA" then
	%><th>ExchangeRate <br> (CAD per USD)</th><%
	end if
	%>
	
	
	<th>Date IN</th>
	<th>Last Modify Date</th>
    <th>Value
	<% 
	if Country ="USA" then
		Response.write " ($USD)"
	else  
		Response.write " ($CAD)"
	end if
	%>
	</th>
	<th>Rolling Total	<% 
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
part = rs("part")
qty = rs("Qty")
po = rs("po")
EnterDate = rs("EnterDate")
LastModify = rs("LastModify")
PPU = rs("PPU") + 0
errormsg = 0


'Create a Query
    SQL2 = "SELECT * FROM Y_HARDWARE_MASTER where PART = '" & part & "' order BY ID DESC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)
	
	if rs2.eof then
		errormsg = 1
		price = 0
		Description = "Error in Description - Value 0"
	else
		If IsNull(PPU) Or Trim(PPU) = 0 Or Trim(PPU) = "0" Or Trim(PPU) = "" Or UCASE(Trim(PPU)) = "NULL" Then
			price = rs2("price")			
		Else
			price = PPU			
		End If		
		description = rs2("Description")
			errormsg = 0
	end if
	rs2.close
	set rs2 = nothing



if errormsg = 1 then
response.write part & "(" & rs("id") & ")" &" not in inventory master <BR>"
end if

'response.write partqty & "<BR>"
If Not IsNull(QTY) Then



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


	TempValue = (qty * Price) / ExchangeRate
	TotalValue = TotalValue + TempValue	

%>

	<tr>
	<td><% response.write part %></td>
    <td><% response.write description %></td>
    <td><% response.write qty %></td>
    <td><% response.write po %></td>
	<td><% response.write Price %></td>
		<% 
		if Country ="USA" then
		%><td><% response.write ExchangeRate %></td><%
		end if
		%>
	
	<td><% response.write EnterDate %></td>
	<td><% response.write LastModify %></td>
    <td>$<% response.write round(tempvalue,2) %></td>
	<td>$<% response.write round(totalvalue,2) %></td>
	

    </tr>
	

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

	
RecordNumbers = RecordNumbers + 1
rs.movenext
loop
rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing


if Country ="USA" then
	cn_SQL.close
	set cn_SQL = nothing
end if

%> 
</Table> 

<%

response.write "<BR><HR>"
response.write "Total Hardware Material $" & round(Totalvalue,2) & "<BR>"
response.write RecordNumbers & " Records"& "<BR>"

%>

       
            
    
</body>
</html>
