<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 10th, by Michael Bernholtz - Entry Form to Update existing items in the QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- Page updated to allow remembering of the fields submitted--> 
<!-- February 2019 - Add USA VIEW - Writes to Seperate USA database -->

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
 
<script type="text/javascript">
    function showForm() {
        var invType = document.getElementById("InventoryType").value;
        if (invType == "QCGlass") {
            document.getElementById("GLASSIDFORM").style.display = "block";
            document.getElementById("SPACERIDFORM").style.display = "none";
			document.getElementById("SEALANTIDFORM").style.display = "none";
			document.getElementById("MISCIDFORM").style.display = "none";
            document.getElementById("Identify").style.display = "block";
			document.getElementById("Amount").style.display = "block";

        }
        if (invType == "QCSpacer") {
            document.getElementById("GLASSIDFORM").style.display = "none";
            document.getElementById("SPACERIDFORM").style.display = "block";
			document.getElementById("SEALANTIDFORM").style.display = "none";
			document.getElementById("MISCIDFORM").style.display = "none";
            document.getElementById("Identify").style.display = "block";
			document.getElementById("Amount").style.display = "block";
        }
        if (invType == "QCSealant") {
            document.getElementById("GLASSIDFORM").style.display = "none";
            document.getElementById("SPACERIDFORM").style.display = "none";
			document.getElementById("SEALANTIDFORM").style.display = "block";
			document.getElementById("MISCIDFORM").style.display = "none";
            document.getElementById("Identify").style.display = "block";
			document.getElementById("Amount").style.display = "none";
        }
		if (invType == "QCMisc") {
            document.getElementById("GLASSIDFORM").style.display = "none";
            document.getElementById("SPACERIDFORM").style.display = "none";
			document.getElementById("SEALANTIDFORM").style.display = "none";
			document.getElementById("MISCIDFORM").style.display = "block";
            document.getElementById("Identify").style.display = "block";
			document.getElementById("Amount").style.display = "block";
        }
    }
</script>
 
 
 
    </head>

<body onload="showForm()">
<%
InventoryType = REQUEST.QueryString("InventoryType")

Select Case InventoryType
	Case "QCGlass"
		QCID1 = REQUEST.QueryString("QCID1")
	
	Case "QCSpacer"
		QCID2 = REQUEST.QueryString("QCID2")	
	
	Case "QCSealant"
		QCID3 = REQUEST.QueryString("QCID3")	
	
	Case "QCMisc"
		QCID4 = REQUEST.QueryString("QCID4")
	
end Select




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
            
              <form id="enter" title="Update QC Inventory" class="panel" name="enter" action="QCInventoryConf.asp" method="GET" target="_webapp" selected="true">
              
			  <h2>Select inventory type, Scan the item, and Input the change in quantity</h2>
			  
              <fieldset>               

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
        
		<div class="row" id='GLASSIDFORM' style="display:block">
			<label>Glass</label>
			<select id='QCID1' name="QCID1" >
 <%
 	if QCID1 <> "" Then

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ItemName FROM QC_MASTER_GLASS WHERE ID=" & QCID1
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
	
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write QCID1
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
		rs2.close   
		set rs2 = nothing
	end if
	
	


		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_GLASS ORDER BY ItemName"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		
			Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("id")
Response.Write "'>"
Response.Write rs("ItemName")
response.write "</option>"

rs.movenext
loop
		rs.close
		set rs = nothing



			
	%>			
			</select>
        </div>

		<div class="row" id='SPACERIDFORM' style="display:None">
			<label>Spacer</label>
			<select id='QCID2' name="QCID2" >
 <%

 	if QCID2 <> "" Then
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ItemName FROM QC_MASTER_SPACER WHERE ID=" & QCID2
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
	
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write QCID2
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
		
			rs2.close   
			set rs2 = nothing
	end if



 Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT Distinct * FROM QC_MASTER_SPACER ORDER BY ItemName"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		
			Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("id")
Response.Write "'>"
Response.Write rs("ItemName")
response.write "</option>"

rs.movenext
loop
rs.close
set rs = nothing

			
	%>			
			</select>
        </div>
		
			<div class="row" id='SEALANTIDFORM' style="display:None">
			<label>Sealant</label>
			<select id='QCID3' name="QCID3" >
 <%

 	if QCID3 <> "" Then
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ItemName FROM QC_MASTER_SEALANT WHERE ID=" & QCID3
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
	
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write QCID3
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
		
		rs2.close   
		set rs2 = nothing
	end if




		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT Distinct * FROM QC_MASTER_SEALANT ORDER BY ItemName"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		
			Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("id")
Response.Write "'>"
Response.Write rs("ItemName")
response.write "</option>"

rs.movenext
loop
rs.close
set rs = nothing

			
	%>			
			</select>
        </div>	
		
		
		<div class="row" id='MISCIDFORM' style="display:None">
			<label>Miscellaneous</label>
			<select id='QCID4' name="QCID4" >
 <%
 
  	if QCID4 <> "" Then
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT ItemName FROM QC_MASTER_MISC WHERE ID=" & QCID4
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
	
		Do While Not rs2.eof
	
			Response.Write "<option value='"
			Response.Write QCID4
			Response.Write "' selected >"
			Response.Write rs2("ItemName")
			response.write "</option>"

		rs2.movenext
		loop
		
		rs2.close   
		set rs2 = nothing
	end if
 
 
 		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT Distinct * FROM QC_MASTER_MISC ORDER BY ItemName"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		
			Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("id")
Response.Write "'>"
Response.Write rs("ItemName")
response.write "</option>"

rs.movenext
loop
rs.close
set rs = nothing

			
	%>			
			</select>
        </div>		
		
		
        <div class="row" id="Identify" Style="display:block">
            <label>Identifier</label>
            <input type="text" name='Identifier' id='Identifier' >
        </div>
		
		<div class="row" id="Amount" Style="display:block">
			<label>Quantity</label>
            <input type="Number" name='Quantity' id='Quantity' Value=1>
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
    
	}
    </script>           
   


<%

On Error Resume Next

DBConnection.close
set DBConnection = nothing
%>

            
</body>
</html>
