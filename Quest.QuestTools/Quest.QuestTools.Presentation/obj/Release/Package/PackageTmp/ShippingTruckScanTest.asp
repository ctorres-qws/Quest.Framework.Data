<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
 <!--#include file="dbpath.asp"-->  
<!-- Created April 2014, by Michael Bernholtz - Scan Existing items to trucks-->
<!-- Dropdown of active Trucks and Scanpad for Shipping items  -->
<!-- Moved to ShippingHome- David Ofir - June 2019-->	

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title> Add to Truck</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script src="sorttable.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
    
</head>

<body>

<%

Call DebugLog(Request.Querystring("barcode"), "-----------------", 0)

Call DebugLog(Request.Querystring("barcode"), "Page - S", 1)

truck = trim(Request.Querystring("truck"))
barcode = trim(Request.Querystring("barcode"))

Scanned = False   ' Flag for an item that was Scanned only shows if Database updated
					' Two ways to Flag (Accessory ID and Barcode)

If Len(barcode) > 0 Then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Call Process(false, true)
		Case c_MODE_HYBRID
			Call Process(false, true)
			Call Process(true, false)
		Case c_MODE_SQL_SERVER
			Call Process(true, true)
	End Select

End If

Function Process(isSQLServer, b_First)

Call DebugLog(Request.Querystring("barcode"), "Process - S ", 2)

Call DebugLog(Request.Querystring("barcode"), "Process DB Connection - S ", 3)

DBOpen DBConnection, isSQLServer

Call DebugLog(Request.Querystring("barcode"), "Process DB Connection - E ", 3)

	'Active Trucks
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = FixSQLCheck("SELECT top 10000 * FROM X_Shipping_Truck_Test WHERE active = TRUE {0} ORDER BY ID DESC", isSQLServer)

	if truck <> "" and truck <> "-" then
		strSQL = Replace(strSQL,"{0}"," AND ID=" & truck & " ",1)
	else
		strSQL = Replace(strSQL,"{0}","",1)
	end if

	'DebugCode(strSQL)
	
	rs.Cursortype = 2
	rs.Locktype = 3
	Call DebugLog(Request.Querystring("barcode"), "Load X_Shipping_Truck - S ", 3)
	rs.Open strSQL, DBConnection
	Call DebugLog(Request.Querystring("barcode"), "Load X_Shipping_Truck - E ", 3)

	' All items (accessory)
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT top 10000 * FROM X_Shipping_Test ORDER BY ID DESC"
	'rs2.Cursortype = GetDBCursorType
	'rs2.Locktype = GetDBLockType
	'rs2.Open strSQL2, DBConnection

	Call DebugLog(Request.Querystring("barcode"), "Load X_Shipping - S ", 3)
	Set rs2 = GetDisconnectedRS(strSQL2, DBConnection)
	Call DebugLog(Request.Querystring("barcode"), "Load X_Shipping - E ", 3)
' Update the Record with the Truck id

Call DebugLog(Request.Querystring("barcode"), "Section 1 - S ", 3)

if truck <> "" and truck <> "-" then
	'rs.filter = "ID = " & truck
	if not rs.eof then
	'Valid Truck
		if barcode <> "" then
		
			if Left(barcode,2) = "GT" AND NOT Left(barcode,3) ="GTM" then
				if len(barcode) <3 then 
					bc = "GT00"
				end if	
				GlassID = Mid(barcode, 3)

				Set rs4 = Server.CreateObject("adodb.recordset")
				strSQL4 = "SELECT top 1000 * FROM Z_GLASSDB ORDER BY ID DESC"
				rs4.Cursortype = 2
				rs4.Locktype = 3
				rs4.Open strSQL4, DBConnection
				rs4.filter = "ID = " & GlassID
				if not rs4.eof then
								jobname = rs4("Job")
								floorname = rs4("Floor")
								tagname = rs4("Tag")
				end if
				rs4.close
				set rs4 = nothing
				WindowValue = "Window"
				DescriptionName = "Service"
			
			else
			
				if Left(barcode,2) = "00"  or Left(barcode,2) = "11" or Left(barcode,2) = "22" then
			
					if Left(barcode,2) = "00"  then
						'Special Code for Other   00JOBFLOOR-Tag:Description
						WindowValue = "Other"
						jobname = Mid(barcode, 3, 3)
						floorname = Mid(Barcode, 6, (inStr(1, barcode, "-", 0) -6))
						tagname = Mid(Barcode, (inStr(1, barcode, "-", 0)+ 1), ((inStr(1, barcode, ":", 0)) - (inStr(1, barcode, "-", 0)))- 1)
						DescriptionName = Mid(barcode, (inStr(1, barcode, ":", 0) + 1), 10)
					end if
					
					if Left(barcode,2) = "11"  then
						'Special Code for HBar   11JOBFLOOR-Tag.Count
						WindowValue = "H-Bar"
						jobname = Mid(barcode, 3, 3)
						floorname = Mid(Barcode, 6, (inStr(1, barcode, "-", 0) -6))
						tagname = Right(Barcode, (len(barcode) - (inStr(1, barcode, "-", 0))))
					'	tagname = Mid(Barcode, (inStr(1, barcode, "-", 0)+ 1), ((inStr(1, barcode, ".", 0)) - (inStr(1, barcode, "-", 0)))- 1)
						DescriptionName = "H-Bar"
					end if
					
					if Left(barcode,2) = "22"  then
						'Special Code for Jamb Receptor   22JOBFLOOR-jamb.COUNT
						WindowValue = "Jamb Receptor"
						jobname = Mid(barcode, 3, 3)
						floorname = Mid(Barcode, 6, (inStr(1, barcode, "-", 0) -6))
						tagname = Right(Barcode, (len(barcode) - (inStr(1, barcode, "-", 0))))
					'	tagname = Mid(Barcode, (inStr(1, barcode, "-", 0)+ 1), ((inStr(1, barcode, ".", 0)) - (inStr(1, barcode, "-", 0)))- 1)
						DescriptionName = "Jamb Receptor"
					end if
					
					
			
				else
					WindowValue = "Window"
					DescriptionName = "Production"
					'Determine Job Floor Tag (must be Unique) for window
					
					' 3 step process to breakdown barcode into Job Floor and tag from Window Barcode
				

					jobname = Left(barcode, 3)
					floorname = Left(barcode, inStr(1, barcode, "-", 0) - 1)
					floorname = Right(floorname, Len(FloorName)-3)
					tagname =  Right(barcode, Len(Barcode)- inStr(1, barcode, "-", 0))
				
'						jobname = Left(barcode, 3)
'						if inStr(1, barcode, "-", 0) = 5 then
'							floorname = Mid(barcode, 4, 1)
'							tagname = Mid(barcode, 6, 8)
'						END IF
'
'						if inStr(1, barcode, "-", 0) = 6 then
'							floorname = Mid(barcode, 4, 2)
'							tagname = Mid(barcode, 7, 8)
'						end if
'
'						if inStr(1, barcode, "-", 0) = 7 then
'							floorname = Mid(barcode, 4, 3)
'							tagname = Mid(barcode, 8, 8)
'						end if	
'						if inStr(1, barcode, "-", 0) = 8 then
'							floorname = Mid(barcode, 4, 4)
'							tagname = Mid(barcode, 9, 8)
'						end if	
				end if
			end if
				' Filter by Barcode - if exists update
				rs2.filter = "job = '" & jobname& "' AND floor = '" & floorname & "' AND tag = '" & tagname & "'"
				if not rs2.eof then
				else 
					Scanned = True
					Shiptime = Hour(Now) & ":" & Minute(Now)
					if hour(now)<= 6 then
						ShipDate = DateAdd("d", -1, Date())
					else
						ShipDate = Date()
					end if
					SQL = "INSERT INTO X_SHIPPING_TEST (BARCODE, JOB, FLOOR, TAG, SHIPDATE, ShipTime, TRUCK, Description, [Window]) VALUES ('" & barcode & "', '" & jobname & "', '" & floorname & "', '" & tagname & "', '" & ShipDate & "', '" & Shiptime & "', '" & truck & "', '" & DescriptionName & "', '" & WindowValue & "' )"				
					'SQL = "INSERT INTO X_SHIPPING_TEST (BARCODE, JOB, FLOOR, TAG, SHIPDATE, ShipTime, TRUCK, Description) VALUES ('" & barcode & "', '" & jobname & "', '" & floorname & "', '" & tagname & "', '" & Date() & "', '" & Shiptime & "', '" & truck & "', '" & DescriptionName & "' )"				

					If b_First Then
					Call DebugLog(Request.Querystring("barcode"), "Check Parent - S ", 4)
					%>
					 <!--#include file="GmailCheckParentTest.asp"--> 
					<%
					Call DebugLog(Request.Querystring("barcode"), "Check Parent - E ", 4)
					End If

					'Get a Record Set
					Set RS_Ship = DBConnection.Execute(SQL)
					'Window is new, add to list
					Set rs4 = Server.CreateObject("adodb.recordset")
					strSQL4 = "SELECT top 1000 * FROM Z_GLASSDB ORDER BY ID DESC"
					rs4.Cursortype = 2
					rs4.Locktype = 3
					rs4.Open strSQL4, DBConnection

					rs4.filter = "JOB = '" & jobname & "' and Floor = '" & floorname & "' and TAG = '" & tagname & "'"
					if not rs4.eof then
						if isDate(rs4.fields("ShipDate")) = False then
							' Record in Z_GLASSDB matches the scanned item and does not have an Output Date
							' Successfully Marked Shipped
							rs4.fields("ShipDate") = Date
							rs4.update
						end if
					end if
					rs4.close
					set rs4 = nothing
			end if
		end if
	end if
	rs.filter = ""
end if

Call DebugLog(Request.Querystring("barcode"), "Section 1 - E ", 3)

Call DebugLog(Request.Querystring("barcode"), "Item Count - S ", 3)

	if not rs.eof then
		rs.movefirst
		Do while Not rs.eof
		Itemcount = 0
			rs2.filter = "truck = '" & rs("ID") & "'"
			if not rs2.eof then
				Do while not rs2.eof
					ItemCount = ItemCount + 1
				rs2.movenext
				loop
				rs("itemcount") = Itemcount
				rs.update
			end if
			rs2.filter = ""
		rs.movenext
		loop
	end if

Call DebugLog(Request.Querystring("barcode"), "Item Count - E ", 3)

Call DebugLog(Request.Querystring("barcode"), "DBCloseAll - S ", 3)

	DbCloseAll

Call DebugLog(Request.Querystring("barcode"), "DBCloseAll - E ", 3)

	Call DebugLog(Request.Querystring("barcode"), "Process - E ", 2)

End Function

%>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
			<a class="button leftButton" type="cancel" href="ShippingHomeTest.HTML" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="AddTruck" title="Scan to Active Truck" class="panel" name="AddTruck" action="ShippingTruckScanTest.asp" method="GET" selected="true">
        
        <h2>Select Active Truck</h2>
       <fieldset>

        <div class="row">
            <label>Truck</label>
            <select name="truck">
<%
	'Active Trucks
Dim rs
Set DBConnection = Server.CreateObject("adodb.connection")
DBOpen DBConnection, isSQLServer

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = FixSQLCheck("SELECT top 1000 * FROM X_Shipping_Truck_Test WHERE active = True ORDER BY ID DESC", isSQLServer)
	'strSQL = FixSQLCheck("SELECT top 1000 * FROM X_Shipping_Truck_Test WHERE active = True {0} ORDER BY ID DESC", isSQLServer)
	
'	if truck <> "" and truck <> "-" then
'		strSQL = Replace(strSQL,"{0}"," AND ID=" & truck & " ",1)
'	else
'		strSQL = Replace(strSQL,"{0}","",1)
'	end if

	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

if truck <> "" and truck <> "-" then
	rs.filter = "ID = " & truck
	activeTruck = ""
	if rs("truckName") <> "" then
		activeTruck = rs("truckName") & " - "
	end if
	activeTruck = activeTruck & rs("Job") & rs("Floor") & "-" & rs("truckNum")
	response.write " <option value = '" & rs("id") & "'>" & activeTruck & "</option>"

end if	
	if truck <> "" and truck <> "-" then
		rs.filter = "ID <> " & truck	
	else
		rs.filter = ""	
	end if
	do while not rs.eof
	Response.Write "<option value = '"
	Response.Write rs("id")
	Response.Write "'>"
	if rs("truckName") <> "" then
		Response.Write rs("truckName") & " - "
	end if
	Response.Write rs("Job") & rs("Floor") & "-" & rs("truckNum")
	rs.movenext
	loop

%>
</select>
	
	    </div>
		
		<div class="row">
                <label>Shipping Item</label>
                <input type="text" name='barcode' id='barcode' >
        </div>
				
        </fieldset>
        <BR>
		<a class="whiteButton" onClick="AddTruck.submit()">Submit</a><BR>
		
		
            <ul id="Profiles" title="Active Trucks" selected="true">
        <%
		if Scanned = True then
			response.write" <li>Scan Successful</li>"
				if truck <> "" and truck <> "-" then
					rs.filter = "ID = " & truck
					activeTruck = rs("truckName") & "-" & rs("Job") & " - Floor " & rs("Floor") 
					response.write" <li>Assigned to Truck Number "& rs("truckNum") & " for : " & activeTruck & "</li>"
					rs.filter = ""
				else
					Response.write "<li> Please Select a Truck for Scanning </li>"
				end if
		else
			if barcode = "" then
			response.write" <li>No Item entered</li>"
			else
			response.write" <li>Item Already Scanned</li>"
			end if
		end if	

		
        %>
			</ul>
            </form>
    <script type="text/javascript">
				  
				 		  
    function callback1(barcode) {
        var barcodeText = "BARCODE:" + barcode;

        document.getElementById('barcode').innerHTML = barcodeText;
        console.log(barcodeText);
        
    }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("barcode");

    textbox.value = barcode;
	AddTruck.submit();

}

        </script>
<%

Call DebugLog(Request.Querystring("barcode"), "DBClose - S", 2)

'DbCloseAll

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
set DBConnection = nothing

Call DebugLog(Request.Querystring("barcode"), "DBClose - E", 2)

Call DebugLog(Request.Querystring("barcode"), "Page - E", 1)

%>	

</body>
</html>