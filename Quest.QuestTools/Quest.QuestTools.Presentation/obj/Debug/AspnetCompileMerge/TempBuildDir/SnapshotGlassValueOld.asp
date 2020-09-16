<!--#include file="dbpath_secondary.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!-- Created for Mahesh Mohanlall and Vanessa Abraham by Michael Bernholtz April 2019-->
<!-- Read and Report Glass information -->
<!-- SnapshotGlassSelect SnapshotGlassValue-->

 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%


Country = Request.Querystring("Country")
if Country ="USA" then
	ReportName = Request.Querystring("ReportName1")
else
	ReportName = Request.Querystring("ReportName2")
end if




ReportMonth = Left(Right(ReportName,6),2)
IF Left(ReportMonth,1) = 0 Then
	ReportMonth = Right(ReportMonth,1)
end if
ReportYear = Right(ReportName,4)
RecordNumbers = 0

' Canada View NASHUA USA view JUPITER - for future reports
if Country ="USA" then
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (Quantity > 0) order by MasterID ASC"
else
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (Quantity > 0) order by MasterID ASC"
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
<h2> Glass Inventory As of  <%Response.write ReportMonth%> / <%Response.write ReportYear%></h2>
<h3> Using Inventory from: <%Response.write ReportName%></h3>
<h3> Viewing <%Response.write Country%></h3>

<table border='1' class='sortable'>
	<tr>
	<th>Item Name</th>
	<th>Size</th>
    <th>Serial Number</th>
    <th>Manufacturer</th>
	<th>Qty</th>
	<th>Pack</th>
	<th>SQFT <br>(Per Pack)</th>
    <th>$/SQFT ($CAD)</th>
	<% 
	if Country ="USA" then
	%><th>ExchangeRate <br> (CAD per USD)</th><%
	end if
	%>
	<th>Value
	<% 
	if Country ="USA" then
		Response.write " ($USD) <br> (Qty * SQFT * Price)/ Exchange"
	else  
		Response.write " ($CAD) <br> (Qty * SQFT * Price)"
	end if
	%>
	</th>
	<th>Entry Date</th>
	<th>Rolling Total
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
masterid = rs("masterid")
SerialNumber = rs("SerialNumber")
Quantity = rs("Quantity")
EntryDate = rs("EntryDate")
errormsg = 0


'Create a Query
    SQL2 = "SELECT * FROM [MS Access;DATABASE=" & "F:\database\QualityControlDB.mdb" & "].[QC_MASTER_GLASS] where ID = " & masterid & " order BY ID DESC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)
	
	if rs2.eof then
		errormsg = 1
		price = 0
		Description = "Error in Description - Value 0"
	else
		itemName  = rs2("ItemName")
		Size  = "" & rs2("Width") & " X " & rs2("Height")
		Manufacturer  = rs2("Manufacturer")
		Pack = rs2("Pieces")
		price = rs2("price")
		SQFT = Round((rs2("Width") * rs2("Height") * rs2("Pieces"))/144,2)
		errormsg = 0
		
	end if
	rs2.close
	set rs2 = nothing



if errormsg = 1 then
response.write part & "(" & rs("id") & ")" &" not in Glass master <BR>"
end if

'response.write partqty & "<BR>"
If Not IsNull(Quantity) Then


	if Country ="USA" then
		PartMonth = Month(EntryDate)
		if Len(PartMonth) = 1 then
			PartMonth = "0" & PartMonth
		end if
		PartYear = Year(EntryDate)
		
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
	
	TempValue = (Price * SQFT * Quantity) / ExchangeRate
	TotalValue = TotalValue + TempValue	

%>

	<tr>
	<td><% response.write itemName %></td>
    <td><% response.write Size %></td>
    <td><% response.write SerialNumber %></td>
    <td><% response.write Manufacturer %></td>
	<td><% response.write Quantity %></td>
	<td><% response.write Pack %></td>
	<td><% response.write SQFT %></td>
	<td><% response.write Price %></td>
	<% 
	if Country ="USA" then
	%><td><% response.write ExchangeRate %></td><%
	end if
	%>
	
    <td>$<% response.write round(tempvalue,2) %></td>
	<td><% response.write EntryDate %></td>
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
response.write "Total Glass Material $" & round(Totalvalue,2) & "<BR>"
response.write RecordNumbers & " Records"& "<BR>"

%>

       
            
    
</body>
</html>
