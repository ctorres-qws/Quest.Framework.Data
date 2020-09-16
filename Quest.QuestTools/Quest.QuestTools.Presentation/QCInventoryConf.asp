<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 6th, by Michael Bernholtz - Confirmation  of input items to QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- February 2019 - Add USA VIEW - Writes to Seperate USA database -->
<!-- USA Glass go to QC_GLASS_USA, Spacer go to QC_Spacer_USA, Sealant go to QC_Sealant_USA, Misc go to QC_Misc_USA-->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add lites quantity when adding to inventory of type Glass
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

<%
IsError = False

InventoryType = REQUEST.QueryString("InventoryType")

Identifier = Trim(REQUEST.QueryString("Identifier"))
if Identifier = "" then
	IsError = True
	Error = "ERROR: No Serial Number / Identifier Scanned"
End if

Quantity = Trim(REQUEST.QueryString("Quantity"))
LitesQty = Trim(REQUEST.QueryString("LitesQty"))
EntryDate = Date()
ResponseR = ""

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

If IsError = False then
Select Case InventoryType
	Case "QCGlass"
		
		if CountryLocation = "USA" then
			strSQL2 = "SELECT * FROM QC_GLASS_USA ORDER BY SerialNumber ASC"
		else
			strSQL2 = "SELECT * FROM QC_GLASS ORDER BY SerialNumber ASC"
		end if
		
		Set rs2 = Server.CreateObject("adodb.recordset")
		rs2.Cursortype = GetDBCursorTypeInsert
		rs2.Locktype = GetDBLockTypeInsert
		rs2.Open strSQL2, DBConnection
		rs2.filter = "SerialNumber = '" & Identifier & "'" 
		
		QCID = REQUEST.QueryString("QCID1")
		if rs2.eof then
			
			'Set Glass Input Statement
			'StrSQL = "INSERT INTO QC_GLASS (MasterID, SerialNumber, EntryDate, Quantity, LitesQty ) VALUES (" & qcid & ", '" & Identifier & "', '" & EntryDate & "', " & Quantity & ", " & LitesQty & ")"
			rs2.AddNew
			rs2.fields("MasterID") = qcid
			rs2.fields("SerialNumber") = Identifier
			rs2.fields("EntryDate") = EntryDate
			rs2.fields("Quantity") = Quantity
			rs2.fields("LitesQty") = LitesQty
			If GetID(isSQLServer,1) <> "" Then rs2.Fields("ID") = GetID(isSQLServer,1)
			rs2.Update
			Call StoreID1(isSQLServer, rs2.Fields("ID"))
			'Get a Record Set
			'Set RS = DBConnection.Execute(strSQL)
			ResponseR = "QC_GLASS Entered"
		else
				IsError = TRUE
				Error = " ERROR: " & Identifier & " Already Scanned."
		end if
		rs2.close
		set rs2 = nothing
		
	Case "QCSpacer"
	
		if CountryLocation = "USA" then
			strSQL2 = "SELECT * FROM QC_SPACER_USA ORDER BY Identifier ASC"
		else
			strSQL2 = "SELECT * FROM QC_SPACER ORDER BY Identifier ASC"
		end if
		Set rs2 = Server.CreateObject("adodb.recordset")
		rs2.Cursortype = GetDBCursorTypeInsert
		rs2.Locktype = GetDBLockTypeInsert
		rs2.Open strSQL2, DBConnection
		rs2.filter = "Identifier = '" & Identifier & "'" 
		
		QCID = REQUEST.QueryString("QCID2")
		if rs2.eof then
		
			'Set Spacer Input Statement
			'StrSQL = "INSERT INTO QC_SPACER (MasterID, Identifier, EntryDate) VALUES (" & qcid & ", '" & Identifier & "', '" & EntryDate & "')"
			rs2.AddNew
			rs2.fields("MasterID") = qcid
			rs2.fields("Identifier") = Identifier
			rs2.fields("EntryDate") = EntryDate
			rs2.fields("Quantity") = Quantity
			If GetID(isSQLServer,1) <> "" Then rs2.Fields("ID") = GetID(isSQLServer,1)
			rs2.Update
			Call StoreID1(isSQLServer, rs2.Fields("ID"))
			'Get a Record Set
			'Set RS = DBConnection.Execute(strSQL)
			ResponseR = "QC_SPACER Entered"
		else
				IsError = TRUE
				Error = " ERROR: " & Identifier & " Already Scanned."
		end if	
		rs2.close	
		set rs2 = nothing
				

	Case "QCSealant"
	
	
		if CountryLocation = "USA" then
			strSQL2 = "SELECT * FROM QC_SEALANT_USA ORDER BY Identifier ASC"
		else
			strSQL2 = "SELECT * FROM QC_SEALANT ORDER BY Identifier ASC"
		end if
		Set rs2 = Server.CreateObject("adodb.recordset")
		rs2.Cursortype = GetDBCursorTypeInsert
		rs2.Locktype = GetDBLockTypeInsert
		rs2.Open strSQL2, DBConnection
		rs2.filter = "Identifier = '" & Identifier & "'" 
		
		QCID = REQUEST.QueryString("QCID3")
		if rs2.eof then
		
			'Set Sealant Input Statement
			'StrSQL = "INSERT INTO QC_SEALANT (MasterID, Identifier, EntryDate) VALUES (" & qcid & ", '" & Identifier & "', '" & EntryDate & "')"
			rs2.AddNew
			rs2.fields("MasterID") = qcid
			rs2.fields("Identifier") = Identifier
			rs2.fields("EntryDate") = EntryDate
			If GetID(isSQLServer,1) <> "" Then rs2.Fields("ID") = GetID(isSQLServer,1)
			rs2.Update
			Call StoreID1(isSQLServer, rs2.Fields("ID"))
			'Get a Record Set
			'Set RS = DBConnection.Execute(strSQL)
			ResponseR = "QC_SEALANT Entered"
		else
				IsError = TRUE
				Error = " ERROR: " & Identifier & " Already Scanned."
		end if	
		rs2.close
		set rs2 = nothing
	
	Case "QCMisc"

		if CountryLocation = "USA" then
			strSQL2 = "SELECT * FROM QC_MISC_USA ORDER BY Identifier ASC"
		else
			strSQL2 = "SELECT * FROM QC_MISC ORDER BY Identifier ASC"
		end if
	
		Set rs2 = Server.CreateObject("adodb.recordset")
		rs2.Cursortype = GetDBCursorTypeInsert
		rs2.Locktype = GetDBLockTypeInsert
		rs2.Open strSQL2, DBConnection
		rs2.filter = "Identifier = '" & Identifier & "'" 

		QCID = REQUEST.QueryString("QCID4")
		if rs2.eof then

		'Set Sealant Input Statement
		'StrSQL = "INSERT INTO QC_MISC (MasterID, Identifier, EntryDate, Quantity ) VALUES (" & qcid & ", '" & Identifier & "', '" & EntryDate & "', " & Quantity & ")"
		rs2.AddNew
		rs2.fields("MasterID") = qcid
		rs2.fields("Identifier") = Identifier
		rs2.fields("EntryDate") = EntryDate
		rs2.fields("Quantity") = Quantity
		If GetID(isSQLServer,1) <> "" Then rs2.Fields("ID") = GetID(isSQLServer,1)
		rs2.Update
		Call StoreID1(isSQLServer, rs2.Fields("ID"))

		'Get a Record Set
			'Set RS = DBConnection.Execute(strSQL)
				ResponseR = "QC_MISC Entered"
		else
				IsError = TRUE
				Error = " ERROR: " & Identifier & " Already Scanned."
		end if
		rs2.close
		set rs2 = nothing
			
	Case Else 
		Response.Write "<h2>Invalid Input</h2>"
				IsError = TRUE
				Error = " ERROR: Invalid Material Type"

End Select
End if	
	

DbCloseAll

End Function

%>
	</head>
<body onload="setTimeout(window.close,5000);">



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


    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
	

If IsError = True then
Response.Write "<li>" & Error & "</li>"
else
	
Select Case InventoryType
	Case "QCGlass"		
		
		Response.Write "<li>" & ResponseR & "</li>"
		Response.Write "<li> Serial #: " & Identifier & "</li>"  	' Serial Number and Quantity for Glass
		Response.Write "<li> Quantity: " & Quantity & "</li>"
		Response.Write "<li> # Of Lites: " & LitesQty & "</li>"		
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		QCID1 = QCID
	Case "QCSpacer"	
	
		Response.Write "<li>" & ResponseR & "</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"  	
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		QCID2 = QCID
	Case "QCSealant"
	
		Response.Write "<li>" & ResponseR & "</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"  	
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		QCID3 = QCID
	Case "QCMisc"		
		
		Response.Write "<li>" & ResponseR & "</li>"
		Response.Write "<li> Identifier: " & Identifier & "</li>"  	
		Response.Write "<li> Quantity: " & Quantity & "</li>"   	' Quantity for MISC
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		QCID4 = QCID
	Case Else 
		Response.Write "<h2>Invalid Input</h2>"

End Select
End if
%>	

                  <a type="button" class="whiteButton"  href="QCInventoryAdd.asp?InventoryType=<%response.write InventoryType%>&QCID1=<%response.write QCID1%>&QCID2=<%response.write QCID2%>&QCID3=<%response.write QCID3%>&QCID4=<%response.write QCID4%>" target="_self">Add Another Item</a>
        

</ul>



</body>
</html>

<% 
On Error Resume Next
DBConnection.close
set DBConnection=nothing
%>

