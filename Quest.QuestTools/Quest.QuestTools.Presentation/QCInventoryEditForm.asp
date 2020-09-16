<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit and Delete Form for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- February 2019 - USA Tables included -->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add lites quantity row for inventory type Glass
-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  
<% 

Dim InventoryType, ItemID
InventoryType = Request.Querystring("InventoryType")
QCID = Request.QueryString("QCid")

Dim Identifier, IdentifierID
'Identifier changes the entry between Serial Number for Glass, Box Number for Spacer, Lot Number for Sealant

	
		Select Case InventoryType
	Case "QCGlass"

		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_Glass_USA"
		else
			strSQL = "SELECT * FROM QC_Glass"
		end if
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & QCID
		
		
	Case "QCSpacer"
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_Spacer_USA"
		else
			strSQL = "SELECT * FROM QC_Spacer"
		end if
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID
		
		
	Case "QCSealant"
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_Sealant_USA"
		else
			strSQL = "SELECT * FROM QC_Sealant"
		end if
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID

	Case "QCMisc"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_Misc_USA"
		else
			strSQL = "SELECT * FROM QC_Misc"
		end if
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID
		
	
	End Select


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCInventoryEdit.asp?InventoryType=<% response.write InventoryType %>" target="_self">Edit Stock</a>
				<a class="button" href="#searchForm" id="clock"></a>
    </div>			
    
    
    <form id="QCedit" title="Edit Stock" class="panel" action="QCInventoryEditConf.asp" name="QCedit"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	

	
	<% 
		Select Case InventoryType
			Case "QCGlass"
			
			
	%>	
	
		<div class="row" >
			<label>Glass Type</label>
			<select id='InvNum' name="InvNum" >
 <%
 	if QCID <> "" Then

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ID, ItemName FROM QC_MASTER_GLASS ORDER BY ItemName"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		rs2.filter = "ID=" & rs("MASTERID")
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write rs("MASTERID")
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
	end if

rs2.filter = ""


			Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("id")
Response.Write "'>"
Response.Write rs2("ItemName")
response.write "</option>"

rs2.movenext
loop


			
	%>			
			</select>
        </div>
	
	
	
	<div class="row">
        <label>Serial #</label>
        <input type="text" name='SerialNumber' id='SerialNumber' value="<% response.write Trim(rs.fields("SerialNumber")) %>" >
    </div>
    <div class="row">
        <label>Quantity</label>
        <input type="Number" name='Quantity' id='Quantity' value="<%response.write rs.fields("Quantity") %>" >
    </div>
	<div class="row">
        <label># Of Lites</label>
        <input type="Number" name='LitesQty' id='LitesQty' value="<%response.write rs.fields("LitesQty") %>" >
    </div>
	<%
			Case "QCSpacer"			
	%>	
	<div class="row" >
			<label>Spacer Type</label>
			<select id='InvNum' name="InvNum" >
 <%
 	if QCID <> "" Then

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ID, ItemName FROM QC_MASTER_SPACER ORDER BY ItemName"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		rs2.filter = "ID=" & rs("MASTERID")
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write rs("MASTERID")
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
	end if

rs2.filter = ""


			Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("id")
Response.Write "'>"
Response.Write rs2("ItemName")
response.write "</option>"

rs2.movenext
loop


			
	%>			
			</select>
    </div>
	
	<div class="row">
        <label>Identifier</label>
        <input type="text" name='Identifier' id='Identifier' value="<% response.write Trim(rs.fields("Identifier")) %>" >
    </div>
		
	<%
			Case "QCSealant"			
	%>	
	
	<div class="row" >
			<label>Sealant Type</label>
			<select id='InvNum' name="InvNum" >
 <%
 	if QCID <> "" Then

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ID, ItemName FROM QC_MASTER_SEALANT ORDER BY ItemName"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		rs2.filter = "ID=" & rs("MASTERID")
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write rs("MASTERID")
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
	end if

rs2.filter = ""


			Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("id")
Response.Write "'>"
Response.Write rs2("ItemName")
response.write "</option>"

rs2.movenext
loop


			
	%>			
			</select>
    </div>

	<div class="row">
        <label>Identifier</label>
        <input type="text" name='Identifier' id='Identifier' value="<% response.write Trim(rs.fields("Identifier")) %>" >
    </div>
		
	<%
			Case "QCMisc"			
	%>	
	
	<div class="row" >
			<label>Misc Type</label>
			<select id='InvNum' name="InvNum" >
 <%
 	if QCID <> "" Then

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ID, ItemName FROM QC_MASTER_MISC ORDER BY ItemName"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		rs2.filter = "ID=" & rs("MASTERID")
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write rs("MASTERID")
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
	end if

rs2.filter = ""


			Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("id")
Response.Write "'>"
Response.Write rs2("ItemName")
response.write "</option>"

rs2.movenext
loop


			
	%>			
			</select>
        </div>
	
	
	
	<div class="row">
        <label>Identifier</label>
        <input type="text" name='Identifier' id='Identifier' value="<% response.write Trim(rs.fields("Identifier")) %>" >
    </div>	
	<div class="row">
        <label>Quantity</label>
        <input type="Number" name='Quantity' id='Quantity' value="<%response.write rs.fields("Quantity") %>" >
    </div>	
		
	<%
				
			End Select
	%>			
		
               
    <div class="row">
        <label>Entry Date</label>
        <input type="text" name='EntryDate' id='EntryDate' value="<%response.write Trim(rs.fields("EntryDate")) %>" >
    </div>        
	    <div class="row">
        <label>Consumed</label>
        <input type="text" name='ConsumeDate' id='ConsumeDate' value="<%response.write Trim(rs.fields("ConsumeDate")) %>" >
    </div> 
            
						<input type="hidden" name='InventoryType' id='InventoryType' value="<%response.write InventoryType %>" />
						<input type="hidden" name='QCID' id='QCID' value="<%response.write QCID %>" />
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="QCedit.action='QCInventoryEditConf.asp'; QCedit.submit()">Submit Changes</a><BR>
		<a class="whiteButton" onClick="QCedit.action='QCInventoryConsumeConf.asp'; QCedit.submit()">Consume Stock</a><BR>
		<a class="redButton" onClick="QCedit.action='QCInventoryDelConf.asp'; QCedit.submit()">Delete Stock</a><BR>
	

            
            </form>
                        
<% 
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%> 
 
</body>
</html>



