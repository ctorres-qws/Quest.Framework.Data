<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created May 16th, by Michael Bernholtz - Scan Form to Consume existing items in the QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- Page updated to allow remembering of the fields submitted--> 
<!-- February 2019 - USA Tables added -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory Update</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
 
 
    </head>

<body onload="showForm()">
<%

IsError = False
' Reset the Variable for locating an errorCode
errorCode = ""




InventoryType = REQUEST.QueryString("InventoryType")
Identifier = trim(Request.QueryString("Identifier"))
currentDate = Date()

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpenQC DBConnection, isSQLServer

'Update  Code Only runs if Identifier is scanned
if Identifier <> "" then


Select Case InventoryType
	Case "QCGlass"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_GLASS_USA ORDER BY ID ASC"
		else
			strSQL = "SELECT * FROM QC_GLASS ORDER BY ID ASC"
		end if
	
	Set rs = Server.CreateObject("adodb.recordset")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "SerialNumber = '" & Identifier & "'"
	
	if rs.eof then
		IsError = True 
		errorCode = Identifier & " not found"
	else
		IF rs("ConsumeDate") <> "" then 
			IsError = True 
			errorCode = Identifier & " already Consumed."
		else
			rs("ConsumeDate") = currentDate
			rs("quantity") = 0
			rs.update
		end if
	end if
	

	
	
	rs.close
	set rs=nothing
	
	
	
	Case "QCSpacer"
		
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_SPACER_USA ORDER BY ID ASC"
		else
			strSQL = "SELECT * FROM QC_SPACER ORDER BY ID ASC"
		end if
		
	Set rs = Server.CreateObject("adodb.recordset")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "Identifier = '" & Identifier & "'"
	
	if rs.eof then
		IsError = True 
		errorCode = Identifier & " not found"
	else
		IF rs("ConsumeDate") <> "" then 
			IsError = True 
			errorCode = Identifier & " already Consumed."
		else
			rs("ConsumeDate") = currentDate
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	
	
	Case "QCSealant"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_SEALANT_USA ORDER BY ID ASC"
		else
			strSQL = "SELECT * FROM QC_SEALANT ORDER BY ID ASC"
		end if

	Set rs = Server.CreateObject("adodb.recordset")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "Identifier = '" & Identifier & "'"
	
	if rs.eof then
		IsError = True 
		errorCode = Identifier & " not found"
	else
		IF rs("ConsumeDate") <> "" then 
			IsError = True 
			errorCode = Identifier & " already Consumed."
		else
			rs("ConsumeDate") = currentDate
			rs.update
		end if
	end if	
	rs.close
	set rs=nothing
	
	Case "QCMisc"
	
		if CountryLocation = "USA" then
			strSQL = "SELECT * FROM QC_MISC_USA ORDER BY ID ASC"
		else
			strSQL = "SELECT * FROM QC_MISC ORDER BY ID ASC"
		end if

	
	Set rs = Server.CreateObject("adodb.recordset")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "Identifier = '" & Identifier & "'"
	
	if rs.eof then
		IsError = True 
		errorCode = Identifier & " not found" 
	else
		IF rs("ConsumeDate") <> "" then 
			IsError = True 
			errorCode = Identifier & " already Consumed."
		else
			rs("ConsumeDate") = currentDate
			rs("quantity") = 0
			rs.update
		end if
	end if	
	rs.close
	set rs=nothing
	
end Select

end if

DbCloseAll

End Function

%>
  <div class="toolbar">
        <h1 id="pageTitle">Add Glass Inventory</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_QC" target="_self">Glass<%response.write HomeSiteSuffix%></a>
    </div> 
   
            
              <form id="enter" title="Consume QC Inventory" class="panel" name="enter" action="QCInventoryScanConsume.asp" method="GET" target="_self" selected="true">
              
			  
			  
			  
			  <h2>Select inventory type and Scan the item to consume</h2>
			  
              <fieldset>               


			<% if Identifier = "" then
				response.write ""
			else %>
			<div class="row">
                <label><% if IsError = False then
				response.write Identifier & " - Consumed" 
				else
				response.write "ERROR: " & errorCode
				end if	

				%></label></div>         

            <% 			
			end if %>
			<div class="row">
			<label>Inventory Type</label>
			<select id="InventoryType" name="InventoryType" onchange="showForm()">
				<%
				If InventoryType <> "" Then
					
			Response.Write "<option value='"
			Response.Write Trim(InventoryType)
			Response.Write "' selected >"
			Response.Write RIGHT(InventoryType, (LEN(InventoryType)-2)) 
			response.write "</option>"
			
				end if

				%>
				
				<option value="QCGlass">Glass</option>
				<option value="QCSpacer">Spacer</option>
				<option value="QCSealant">Sealant</option>
				<option value="QCMisc">Misc</option>
			</select>
			</div>

        <div class="row" id="Identify" Style="display:block">
            <label>Identifier</label>
            <input type="text" name='Identifier' id='Identifier' >
        </div>
		
            
        <a class="whiteButton" href="javascript:enter.submit()">Submit</a>

</fieldset>


            
            </form>
   <script type="text/javascript">
				  
				 		  
    function callback1(barcode) 
	{
		var barcodeText = "BARCODE:" + barcode;

		document.getElementById('barcode').innerHTML = barcodeText;
		console.log(barcodeText);
        
    }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) 
	{
		var textbox = document.getElementById("Identifier");
 
        textbox.value = barcode;
		enter.submit();
    
	}
    </script>

<%
'rs.close
'set rs = nothing
'rs2.close   
'set rs2 = nothing
'DBConnection.close
'set DBConnection = nothing
%>


</body>
</html>
