<!--#include file="dbpath.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%


RecordNumbers = 0


	strSQL = "SELECT * FROM Y_HARDWARE WHERE (WAREHOUSE = 'PRODUCTION' AND MOVEDJOB LIKE '%Texas%' ) order by PART ASC"



Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Aluminum price
alumprice = AlumPrice

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<h2> Hardware Inventory Moved to Texas</h2>
<h3> Using Inventory from: Main Hardware</h3>

<table border='1' class='sortable'>
	<tr>
	<th>Part</th>
	<th>Description</th>
    <th>Qty</th>
    <th>PO</th>
	<th>Price</th>
	<th>Date IN</th>
	<th>Last Modify Date</th>
    <th>Value</th>
	<th>Rolling Total</th>

    </tr>

<%

rs.movefirst
do while not rs.eof
part = rs("part")
qty = rs("Qty")
po = rs("po")
EnterDate = rs("EnterDate")
LastModify = rs("LastModify")
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

	TempValue = qty * Price
	TotalValue = TotalValue + TempValue	

%>

	<tr>
	<td><% response.write part %></td>
    <td><% response.write description %></td>
    <td><% response.write qty %></td>
    <td><% response.write po %></td>
	<td><% response.write Price %></td>
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

%> 
</Table> 

<%

response.write "<BR><HR>"
response.write "Total Hardware Material $" & round(Totalvalue,2) & "<BR>"
response.write RecordNumbers & " Records"& "<BR>"

%>

       
            
    
</body>
</html>
