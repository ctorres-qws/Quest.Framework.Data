<%@ Language=VBScript %>
<%

ReportName = Request.Querystring("ReportName")
AlumPrice = Request.Querystring("AlumPrice")

dim Cn,Rs
set Cn=server.createobject("ADODB.connection")
set Cn2=server.createobject("ADODB.connection")
set Rs=server.createobject("ADODB.recordset")
set Rs2=server.createobject("ADODB.recordset")
Cn.open "provider=microsoft.jet.oledb.4.0;data source=" & server.mappath("database2/quest.mdb")
Cn2.open "provider=microsoft.jet.oledb.4.0;data source=" & server.mappath("database2/InventoryReports.mdb")




Rs.open "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'NASHUA') order by Part ASC" ,Cn2,1,3

Rs2.open "SELECT * FROM Y_HARDWARE_MASTER order BY ID DESC" ,Cn,1,3


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=QCInventoryReport.xls"
	if Rs.eof <> true then
	

' Introductory Information	
	response.write "<table>"
	response.write "<tr><td> Harware Inventory</td></tr>"
	response.write "<tr><td> Using Inventory from: " & ReportName & "</td></tr>"
	
	response.write "</table>"
	

' Main Information from Inv SnapShot	
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<th>Part</th>"
	response.write "<th>Description</th>"
	response.write "<th>Type</th>"
	response.write "<th>Supplier</th>"
	response.write "<th>Price per Unit</th>"
	response.write "<th>Qty</th>"
	response.write "<th>Value</th></tr>"

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

	
	response.write "</table>"
	
	' Total Value of Inventory

	response.write "<table>"
	response.write "<tr><td>Grand Total = $" & formatcurrency(TotalValue) & "</td></tr>"
	response.write "</table>"

	
Rs.close
set Rs=nothing
	Rs2.close
	set Rs2 = nothing


Cn.close
set Cn=nothing
Cn2.close
set Cn2=nothing
%>
