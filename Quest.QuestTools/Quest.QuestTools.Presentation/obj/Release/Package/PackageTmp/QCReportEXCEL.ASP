<%@ Language=VBScript %>
<!--#include file="QCdbpath.asp"-->
<%
jobname = request.QueryString("jobname")
fl = request.QueryString("fl")
mode = request.QueryString("mode")
supplier = request.QueryString("supplier")
dim Cn,Rs,Rs2,Rs3,Rs4
set Cn=server.createobject("ADODB.connection")
set Rs=server.createobject("ADODB.recordset")
set Rs2=server.createobject("ADODB.recordset")
set Rs3=server.createobject("ADODB.recordset")
set Rs4=server.createobject("ADODB.recordset")
Cn.open GetConnectionStrQC(b_SQL_Server)

if CountryLocation = "USA" then
	Rs.open  "SELECT MG.ItemName, MG.Manufacturer, MG.Pieces, MG.Width, MG.Height, MG.Price, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate FROM QC_GLASS_USA AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs2.open "SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.ConsumeDate FROM QC_SPACER_USA AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs3.open "SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.Printed , SE.ConsumeDate FROM QC_SEALANT_USA AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs4.open "SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate FROM QC_MISC_USA AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID ORDER BY MASTERID ASC",Cn,1,3

else
	Rs.open  "SELECT MG.ItemName, MG.Manufacturer, MG.Pieces, MG.Width, MG.Height, MG.Price, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate FROM QC_GLASS AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs2.open "SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.ConsumeDate FROM QC_SPACER AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs3.open "SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.Printed , SE.ConsumeDate FROM QC_SEALANT AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID ORDER BY MASTERID ASC",Cn,1,3
	Rs4.open "SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate FROM QC_MISC AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID ORDER BY MASTERID ASC",Cn,1,3
end if



'Glass
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=InventoryReport.xls"
	if Rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr><th>Item Name</th><th>Size</th><th>Serial Number</th><th>Manufacturer</th><th>Quantity</th><th>Pack</th><th> Square Foot</th><th>$/Sqft</th><th>Value</th><th>Entry Date</th><th>Consumed Date</th></tr>"
	while not Rs.eof
	
		response.write "<tr><td>" & RS("ItemName") &"</td><td>" & RS("Width") & " X " & RS("Height")& "</td><td>" & RS("SerialNumber") & "</td><td>" & RS("Manufacturer") & "</td><td>" & RS("Quantity") & "</td><td>" & RS("Pieces") & "</td><td>" & Round(RS("Width")*RS("Height")*RS("Pieces")/144,2) & "</td><td>$" & RS("Price") & "</td><td>$" & Round(RS("Width")*RS("Height")*RS("Pieces")*RS("Price") /144,2) & "</td><td>" & RS("EntryDate") & "</td><td>" & RS("ConsumeDate") & "</td></tr>"

	Rs.movenext
	wend
	response.write "</table>"
	end if
	
	
'Spacer
if Rs2.eof <> true then
	response.write "<table border=1>"
	response.write "<tr><th>Item Name</th><th>Serial Number</th><th>Manufacturer</th><th>Quantity</th><th>Entry Date</th><th>Consumed Date</th></tr>"

	while not Rs2.eof
	
		response.write "<tr><td>" & rs2("ItemName") &"</td><td>" & rs2("Identifier") & "</td><td>" & rs2("Manufacturer") & "</td><td>" & rs2("EntryDate") & "</td><td>" & rs2("ConsumeDate") & "</td></tr>"
	Rs2.movenext
	wend
	response.write "</table>"
	end if
'Sealant

if Rs3.eof <> true then
	response.write "<table border=1>"
	response.write "<tr><th>Item Name</th><th>Serial Number</th><th>Manufacturer</th><th>Quantity</th><th>Entry Date</th><th>Consumed Date</th></tr>"

	while not Rs3.eof
	
		response.write "<tr><td>" & rs3("ItemName") &"</td><td>" & rs3("Identifier") & "</td><td>" & rs3("Manufacturer") & "</td><td>" & rs3("EntryDate") & "</td><td>" & rs3("ConsumeDate") & "</td></tr>"

	Rs3.movenext
	wend
	response.write "</table>"
	end if
'Misc

if Rs4.eof <> true then
	response.write "<table border=1>"
	response.write "<tr><th>Item Name</th><th>Serial Number</th><th>Manufacturer</th><th>Quantity</th><th>Entry Date</th><th>Consumed Date</th></tr>"

	while not Rs4.eof
	
		response.write "<tr><td>" & rs4("ItemName") &"</td><td>" & rs4("Identifier") & "</td><td>" & rs4("Manufacturer") & "</td><td>" & rs4("Quantity") & "</td><td>" & rs4("EntryDate") & "</td><td>" & rs4("ConsumeDate") & "</td></tr>"

	Rs4.movenext
	wend
	response.write "</table>"
	end if	

Rs.close	
Rs2.close	
Rs3.close	
Rs4.close	
	
set Rs = nothing
set Rs2 = nothing
set Rs3 = nothing
set Rs4 = nothing
Cn.close
Set Cn = nothing
DBConnection.close
Set DBConnection = nothing
%>
