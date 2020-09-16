<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- Updated February 26th to include Consumed and the ability to clear consumption and reactivate -->
<!-- February 2019 - Added USA Tables -->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add lites quantity for inventory type Glass
-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Edit QC Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>

<%
Passkey = "DLL"
Password = UCASE(TRIM(Request.Form("pwd")))

InventoryType = REQUEST.QueryString("InventoryType")
qcid = request.querystring("qcid")

Identifier = Request.Querystring("Identifier")
SerialNumber = Request.Querystring("SerialNumber")
Quantity = Request.Querystring("Quantity")
if Quantity = "" then
	Quantity = 0
end if
LitesQty = Request.Querystring("LitesQty")
if LitesQty = "" then
	LitesQty = 0
end if
EntryDate = Request.Querystring("EntryDate")
ConsumeDate = Request.Querystring("ConsumeDate")
InvNum = Request.Querystring("InvNum")
%>

	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCInventoryEditForm.asp?InventoryType=<% response.write InventoryType %>&QCID=<% response.write qcid %>" target="_self">Edit Stock</a>
    </div>
<%

if Password = Passkey then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			'Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

End If

Function Process(isSQLServer)

DBOpenQC DBConnection, isSQLServer

'Automatically Set Consumed on a Quantity of 0 
if Quantity = 0 and (InventoryType="QCGlass" or InventoryType = "QCMisc") then
	if ConsumeDate = "" then
		ConsumeDate = Date()
	end if
	ConsumeDateEntry = "'" & ConsumeDate & "'"
else 
	ConsumeDateEntry = ""
end if
' Setting a Null value is tricky for Datetime format and requires you to remove the "" from the SQL String, so the if checks for Null and if not null it adds the "" back in Manually
if ConsumeDateEntry = "" then
	ConsumeDateEntry = "NULL"
else
	ConsumeDateEntry = "'" & ConsumeDate & "'"
end if

	
		Select Case InventoryType
	Case "QCGlass"
	
			'Set Glass Inventory Update Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_GLASS_USA SET SerialNumber='"& SerialNumber & "', Quantity='" & Quantity & "', LitesQty='" & LitesQty & "', EntryDate='" & EntryDate & "', ConsumeDate= " & ConsumeDateEntry & ", MASTERID='" & InvNum & "'  WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_GLASS SET SerialNumber='"& SerialNumber & "', Quantity='" & Quantity & "', LitesQty='" & LitesQty & "', EntryDate='" & EntryDate & "', ConsumeDate= " & ConsumeDateEntry & ", MASTERID='" & InvNum & "'  WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		
	Case "QCSpacer"
	
			'Set Glass Inventory Update Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_SPACER_USA SET Identifier='"& Identifier & "', EntryDate='" & EntryDate & "', ConsumeDate=" & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_SPACER SET Identifier='"& Identifier & "', EntryDate='" & EntryDate & "', ConsumeDate=" & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)

	Case "QCSealant"
	
			'Set Sealant Inventory Update Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_SEALANT_USA SET Identifier='"& Identifier & "', EntryDate='" & EntryDate & "', ConsumeDate=" & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_SEALANT SET Identifier='"& Identifier & "', EntryDate='" & EntryDate & "', ConsumeDate=" & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	
	Case "QCMisc"
	
		'Set Miscellaneous Inventory Update Statement
		if CountryLocation = "USA" then
			StrSQL = "UPDATE QC_Misc_USA SET Identifier='"& Identifier & "', Quantity='" & Quantity & "', EntryDate='" & EntryDate & "', ConsumeDate= " & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
		else
			StrSQL = "UPDATE QC_Misc SET Identifier='"& Identifier & "', Quantity='" & Quantity & "', EntryDate='" & EntryDate & "', ConsumeDate= " & ConsumeDateEntry & ", MASTERID='" & InvNum & "'    WHERE ID = " & QCID
		end if
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)
	
	End Select


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

DbCloseAll

End Function

%>


 <% 
if Password = Passkey then
%>          
    

		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
             	

<form id="conf" title="Edit Stock" class="panel" name="conf" action="<%response.write Homesite%>#_QC" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>

<%
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="QCInventoryEditConf.asp?InventoryType=<%response.write InventoryType%>&QCID=<%response.write QCID%>&Identifier=<%response.write Identifier%>&SerialNumber=<%response.write SerialNumber%>&Quantity=<%response.write Quantity%>&LitesQty=<%response.write LitesQty%>&EntryDate=<%response.write EntryDate%>&ConsumeDate=<%response.write ConsumeDate%>&InvNum=<%response.write InvNum%>" method="post" target="_self" selected="True">



<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>




<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
end if
	
if Password = Passkey then
%>
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
Select Case InventoryType
	Case "QCGlass"
		Response.Write "<li>QC Inventory GLASS Edited:</li>"
		Response.Write "<li> Serial Number: " & SerialNumber & "</li>"
		Response.Write "<li> Quantity: " & Quantity & "</li>"
		Response.Write "<li> # Of Lites: " & LitesQty & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		ConsumeDateEntry = Replace(ConsumeDateEntry,"'","")
		ConsumeDateEntry = Replace(ConsumeDateEntry,"NULL","")
		Response.Write "<li> Date of Consumption " & ConsumeDateEntry & "</li>"
		
	Case "QCSpacer"
		Response.Write "<li>QC Inventory SPACER Edited:</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"	
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		ConsumeDateEntry = Replace(ConsumeDateEntry,"'","")
		ConsumeDateEntry = Replace(ConsumeDateEntry,"NULL","")
		Response.Write "<li> Date of Consumption " & ConsumeDateEntry & "</li>"	
		
	Case "QCSealant"
		Response.Write "<li>QC Inventory SEALANT Edited:</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"	
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		ConsumeDateEntry = Replace(ConsumeDateEntry,"'","")
		ConsumeDateEntry = Replace(ConsumeDateEntry,"NULL","")
		Response.Write "<li> Date of Consumption " & ConsumeDateEntry & "</li>"
		
	Case "QCMisc"
		Response.Write "<li>QC Inventory MISCELLANEOUS Edited:</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"
		Response.Write "<li> Quantity: " & Quantity & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		ConsumeDateEntry = Replace(ConsumeDateEntry,"'","")
		ConsumeDateEntry = Replace(ConsumeDateEntry,"NULL","")
		Response.Write "<li> Date of Consumption " & ConsumeDateEntry & "</li>"		
		
End Select	
%>

        <BR>
       
         <a class="whiteButton" href="QCInventoryEdit.asp?InventoryType=<% response.write InventoryType %>" target="_self"> Back</a>

            </form>

     <%
	
end if
%>

<%
On Error Resume Next
DBConnection.close
set DBConnection=nothing
%>

</body>
</html>

