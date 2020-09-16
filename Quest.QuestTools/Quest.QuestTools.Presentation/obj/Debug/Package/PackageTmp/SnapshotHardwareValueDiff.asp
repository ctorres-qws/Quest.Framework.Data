<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%

'ReportName = Request.Querystring("ReportName")
ReportName = "Y_INV052019"
ReportName2 = "Y_INV062019"
Country = Request.Querystring("Country")
ReportMonth = Left(Right(ReportName,6),2)
IF Left(ReportMonth,1) = 0 Then
	ReportMonth = Right(ReportMonth,1)
end if
ReportYear = Right(ReportName,4)
RecordNumbers = 0


' Canada View NASHUA USA view JUPITER - for future reports
if Country ="USA" then
	strSQL = "SELECT * FROM [" & ReportName & "] as a, [" & ReportName2 & "] as b,  WHERE (WAREHOUSE = 'JUPITER') order by PART ASC"
else
	strSQL = "SELECT a.part as parta, a.qty as qtya, a.po as poa, a.enterdate as enterdatea FROM [" & ReportName & "] as a, [" & ReportName2 & "] as b WHERE a.enterDate = b.enterdate and a.part = b.part and a.PO = b.PO and a.Warehouse = b.warehouse and (a.WAREHOUSE = 'NASHUA') order by a.PART ASC"
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
part = rs("parta")
qty = rs("Qtya")
po = rs("poa")
EnterDate = rs("[Y_INV052019]!EnterDatea")
'LastModify = rs("LastModifya")
'partb = rs("b.part")
'qtyb = rs("b.Qty")
'pob = rs("b.po")
'EnterDateb = rs("b.EnterDate")
'LastModifyb = rs("b.LastModify")

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
		price = rs2("price")
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
