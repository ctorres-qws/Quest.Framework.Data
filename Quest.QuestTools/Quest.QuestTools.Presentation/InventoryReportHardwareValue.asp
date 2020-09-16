<!--#include file="dbpath_Quest_InventoryReports.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%

ReportName = Request.Querystring("ReportName")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'NASHUA') order by Part ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection2


'Aluminum price
alumprice = AlumPrice

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<h2> Hardware Inventory from  Nashua </h2>
<h3> Using Inventory from: <%Response.write ReportName%></h3>

<a href="InventoryReportHardwareExcel.asp?ReportName=<%response.write ReportName%>" target="_self"><b>Download Excel Copy</a>
<table border='1' class='sortable'>
	<tr>
	<th>Part</th>
	<th>Description</th>
	<th>Type</th>
	<th>Supplier</th>
	<th>Price per Unit</th>
	<th>Qty</th>
	<th>Value</th>
    </tr>

<%

rs.movefirst
TotalValue = 0
do while not rs.eof
invpart = rs("part")
partqty = rs("Qty")
errormsg = 0


'Create a Query
    SQL2 = "SELECT * FROM Y_HARDWARE_MASTER where PART = '" & invpart & "' order BY ID DESC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)
	
	if rs2.eof then
		errormsg = 1
	else
		HardwarePrice = rs2("Price")
		HardwareType = rs2("type")
		HardwareSupplier = rs2("supplier")
		HardwareDescription = rs2("description")
		errormsg = 0
	end if
	rs2.close
	set rs2 = nothing



if errormsg = 1 then
response.write invpart & " not in Hardware Inventory Master <BR>"
end if

If Not IsNull(PARTQTY) Then

%>

  <tr><td><% response.write invpart %></td>
    <td><% response.write hardwareDescription %></td>
    <td><% response.write hardwareType %></td>
    <td><% response.write hardwareSupplier %></td>
	 <td><% response.write Hardwareprice %></td>
	<td><% response.write partQTY %></td>

    <td> 
	<% TotalValue = TotalValue + Round(HardwarePrice * partQty,2)%>
	<% response.write formatcurrency(Round(HardwarePrice * partQty,2)) %> </td>

    </tr>

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

	

rs.movenext
loop
rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing

%> </Table> <%



response.write "Grand Total = $" & formatcurrency(TotalValue)

%>

       
            
    
</body>
</html>
