<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit and Delete Form for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add lites quantity column for inventory type Glass
-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<script>
function RedirectSort(InventoryType, QCID){

	window.location.href = "QCInventoryEditForm.asp?InventoryType="+InventoryType+"&QCID="+QCID
}
</script>

<%

Dim InventoryType
InventoryType = Trim(Request.Querystring("InventoryType"))

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">QC Inventory Edit</h1>
        <a class="button leftButton" type="cancel" href="QCInventorySelect.asp" target="_self">Type Select</a>
   </div>

<%
Select Case InventoryType
	Case "QCGlass"

		
		if CountryLocation = "USA" then
			strSQL = "SELECT MG.ItemName, MG.Manufacturer, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate, G.printed, G.Id, G.LitesQty FROM QC_GLASS_USA AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID ORDER BY SERIALNUMBER ASC"
		else
			strSQL = "SELECT MG.ItemName, MG.Manufacturer, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate, G.printed, G.Id, G.LitesQty FROM QC_GLASS AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID ORDER BY SERIALNUMBER ASC"
		end if
	
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3

		rs.Open strSQL, DBConnection
		Response.Write " <ul id='Glass' title=' QC Glass Inventory' selected='true'> "
		Response.write "<li class='group'>Select Glass to Manage </li>"

		Response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		Response.write "<li><table border='1' class='sortable' width='95%'><tr><th  width='25%'>Item Name</th><th  width='10%'>Serial Number</th><th  width='10%'>Manufacturer</th><th  width='10%'># of Packs</th><th  width='12.5%'># of Lites/Pack</th><th  width='10%'>EntryDate</th><th  width='12.5%'>Consumed Date</th></tr>"

		Do While not rs.eof

			Response.write "<tr><td>" & trim(RS("ItemName")) &"</td><td>" & trim(RS("SerialNumber")) & "</td><td>" & trim(RS("Manufacturer")) & "</td><td>" & trim(RS("Quantity")) & "</td><td>" & trim(RS("LitesQty")) & "</td><td>" & trim(RS("EntryDate")) & "</td><td>" & trim(RS("ConsumeDate")) & "</td>"
			'THis line does not work in Classic ASP - it plays havoc with the brackets so it is back in HTML
			'The Sorting function conflicts with the variables when sorted, so this is the work around
%>
			<td><input type ='submit' value = 'Manage Glass' onclick="RedirectSort('<%response.write trim(InventoryType)%>','<%response.write trim(RS.fields("ID"))%>')"</td>
<%
			Response.write "</tr>"

			rs.movenext
		Loop
		Response.write "</table></li></ul>"
		rs.close
		set rs = nothing

	Case "QCSpacer"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.ConsumeDate, SP.printed, SP.Id FROM QC_SPACER_USA AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID ORDER BY IDENTIFIER ASC"
		else
			strSQL = "SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.ConsumeDate, SP.printed, SP.Id FROM QC_SPACER AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID ORDER BY IDENTIFIER ASC"
		end if
	
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		Response.Write " <ul id='Spacer' title=' QC Spacer Inventory' selected='true'> "
		Response.write "<li class='group'>Select Spacer to Manage </li>"

		Response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		Response.write "<li><table border='1' class='sortable' width='95%'><tr><th  width='30%'>Item Name</th><th  width='25%'>Identifier</th><th  width='12.5%'>Manufacturer</th><th  width='10%'>EntryDate</th><th  width='12.5%'>Consumed Date</th></tr>"

		Do While Not rs.eof

			Response.write "<tr><td>" & trim(RS("ItemName")) &"</td><td>" & trim(RS("Identifier")) & "</td><td>" & trim(RS("Manufacturer")) & "</td><td>" & trim(RS("EntryDate")) & "</td><td>" & trim(RS("ConsumeDate")) & "</td>"
			'THis line does not work in Classic ASP - it plays havoc with the brackets so it is back in HTML
			'The Sorting function conflicts with the variables when sorted, so this is the work around
%>
			<td><input type ='submit' value = 'Manage Spacer' onclick="RedirectSort('<%response.write trim(InventoryType)%>','<%response.write trim(RS.fields("ID"))%>')"</td>
<%
			Response.write "</tr>"

			rs.movenext
		Loop
		Response.write "</table></li></ul>"
		rs.close
		set rs = nothing
		
	Case "QCSealant"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.ConsumeDate, SE.printed, SE.Id FROM QC_SEALANT_USA AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID ORDER BY IDENTIFIER ASC"
		else
			strSQL = "SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.ConsumeDate, SE.printed, SE.Id FROM QC_SEALANT AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID ORDER BY IDENTIFIER ASC"
		end if
	
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		Response.Write " <ul id='Sealant' title=' QC Sealant Inventory' selected='true'> "
		Response.write "<li class='group'>Select Sealant to Manage </li>"

		Response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		Response.write "<li><table border='1' class='sortable' width='95%'><tr><th  width='30%'>Item Name</th><th  width='25%'>Identifier</th><th  width='12.5%'>Manufacturer</th><th  width='10%'>EntryDate</th><th  width='12.5%'>Consumed Date</th></tr>"

		Do While not rs.eof

			Response.write "<tr><td>" & trim(RS("ItemName")) &"</td><td>" & trim(RS("Identifier")) & "</td><td>" & trim(RS("Manufacturer")) & "</td><td>" & trim(RS("EntryDate")) & "</td><td>" & trim(RS("ConsumeDate")) & "</td>"
			'THis line does not work in Classic ASP - it plays havoc with the brackets so it is back in HTML
			'The Sorting function conflicts with the variables when sorted, so this is the work around
			%>
			<td><input type ='submit' value = 'Manage Sealant' onclick="RedirectSort('<%response.write trim(InventoryType)%>','<%response.write trim(RS.fields("ID"))%>')"</td>
			<%
			Response.write "</tr>"

			rs.movenext
		Loop
		Response.write "</table></li></ul>"
		rs.close
		set rs = nothing

	Case "QCMisc"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate, M.printed, M.Id FROM QC_MISC_USA AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID ORDER BY IDENTIFIER ASC"
		else
			strSQL = "SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate, M.printed, M.Id FROM QC_MISC AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID ORDER BY IDENTIFIER ASC"
		end if
	
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		Response.Write " <ul id='Misc' title=' QC Miscellaneous Inventory' selected='true'> "
		Response.write "<li class='group'>Select Item to Manage </li>"

		Response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		Response.write "<li><table border='1' class='sortable' width='95%'><tr><th  width='25%'>Item Name</th><th  width='20%'>Identifier</th><th  width='12.5%'>Manufacturer</th><th  width='10%'>Quantity</th><th  width='10%'>EntryDate</th><th  width='12.5%'>Consumed Date</th></tr>"

		Do While Not rs.eof

			Response.write "<tr><td>" & trim(RS("ItemName")) &"</td><td>" & trim(RS("Identifier")) & "</td><td>" & trim(RS("Manufacturer")) & "</td><td>" & trim(RS("Quantity")) & "</td><td>" & trim(RS("EntryDate")) & "</td><td>" & trim(RS("ConsumeDate")) & "</td>"
			'THis line does not work in Classic ASP - it plays havoc with the brackets so it is back in HTML
			'The Sorting function conflicts with the variables when sorted, so this is the work around
			%>
			<td><input type ='submit' value = 'Manage Misc Item' onclick="RedirectSort('<%response.write trim(InventoryType)%>','<%response.write trim(RS.fields("ID"))%>')"</td>
			<%
			Response.write "</tr>"

			rs.movenext
		Loop
		Response.write "</table></li></ul>"
		rs.close
		set rs = nothing

	Case Else
			
		Response.Write " <ul id='Invalid' title=' Invalid Choice' selected='true'> "
		Response.write "<li><h2> Invalid Selection </h2></li>"
		Response.Write " </ul>"
	
	End Select

on error resume next
DBConnection.Close
Set DBConnection = nothing	
	
%>

</body>
</html>
