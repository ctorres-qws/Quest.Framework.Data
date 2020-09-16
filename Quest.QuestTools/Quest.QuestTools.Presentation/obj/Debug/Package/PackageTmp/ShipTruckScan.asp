<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
 <!--#include file="dbpath.asp"-->  
<!-- Created April 2014, by Michael Bernholtz - Scan Existing items to trucks-->
<!-- Dropdown of active Trucks and Scanpad for Shipping items  -->
<!-- Moved to ShipHome- David Ofir - June 2019-->	

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
dockNum = trim(Request.Querystring("dockNum"))
%>
  <title> Add to Truck</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
 <meta http-equiv="refresh" content="15;url=http://172.18.13.31:8081/ShipTruckScan.asp?dockNum=<%response.write dockNum%>">
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script src="sorttable.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
    
</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">Dock: <%Response.write dockNum%></h1>
    </div>

<%

truck = trim(Request.Querystring("truck"))
barcode = trim(Request.Querystring("barcode"))

Scanned = False   ' Flag for an item that was Scanned only shows if Database updated
					' Two ways to Flag (Accessory ID and Barcode)
errorDesc = "Item Already Scanned"

If Len(barcode) > 0 Then

gi_mode = c_MODE_ACCESS
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
	strSQL = FixSQLCheck("SELECT top 100 * FROM X_SHIP_TRUCK WHERE active = TRUE {0} ORDER BY ID DESC", isSQLServer)

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
	strSQL2 = "SELECT top 10000 * FROM X_SHIP WHERE [DELETED] = FALSE ORDER BY ID DESC"
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
								jobname = UCASE(rs4("Job"))
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
						jobname = UCASE(Mid(barcode, 3, 3))
						floorname = Mid(Barcode, 6, (inStr(1, barcode, "-", 0) -6))
						tagname = Mid(Barcode, (inStr(1, barcode, "-", 0)+ 1), ((inStr(1, barcode, ":", 0)) - (inStr(1, barcode, "-", 0)))- 1)
						DescriptionName = Mid(barcode, (inStr(1, barcode, ":", 0) + 1), 10)
					end if
					
					if Left(barcode,2) = "11"  then
						'Special Code for HBar   11JOBFLOOR-Tag.Count
						WindowValue = "H-Bar"
						jobname = UCASE(Mid(barcode, 3, 3))
						floorname = Mid(Barcode, 6, (inStr(1, barcode, "-", 0) -6))
						tagname = Right(Barcode, (len(barcode) - (inStr(1, barcode, "-", 0))))
					'	tagname = Mid(Barcode, (inStr(1, barcode, "-", 0)+ 1), ((inStr(1, barcode, ".", 0)) - (inStr(1, barcode, "-", 0)))- 1)
						DescriptionName = "H-Bar"
					end if
					
					if Left(barcode,2) = "22"  then
						'Special Code for Jamb Receptor   22JOBFLOOR-jamb.COUNT
						WindowValue = "Jamb Receptor"
						jobname = UCASE(Mid(barcode, 3, 3))
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
				

					jobname = UCASE(Left(barcode, 3))
					floorname = Left(barcode, inStr(1, barcode, "-", 0) - 1)
					floorname = Right(floorname, Len(FloorName)-3)
					tagname =  Right(barcode, Len(Barcode)- inStr(1, barcode, "-", 0))
				
				end if
			end if
				' Filter by Barcode - if exists update
				rs2.filter = "job = '" & jobname& "' AND floor = '" & floorname & "' AND tag = '" & tagname & "'"
				if not rs2.eof then
				else 
					
					'--------Check Truck's Acceptable List
					AcceptableList = UCASE(rs("sList"))
					JobFloor = UCASE(jobname) & UCASE(floorname)
					if Instr(AcceptableList, JobFloor)<1 then
						Scanned = False
						errorDesc = "Not on Assigned Job/Floor    DO NOT PUT ON TRUCK     Please See Supervisor"
					else
					'--------Check Truck's Acceptable List
						if (Instr(Tagname, "#")>0) OR (Instr(Tagname, ".")>0) then
							Scanned = False
							errorDesc = "Not Entered- Please Scan Shipping Label"
						else
							'-------- Scan item to shipped

							Shiptime = Time
							if hour(now)<= 6 then
								ShipDate = DateAdd("d", -1, Date())
							else
								ShipDate = Date()
							end if
							SQL = "INSERT INTO X_SHIP (BARCODE, JOB, FLOOR, TAG, SHIPDATE, ShipTime, TRUCK, Description, [Window]) VALUES ('" & barcode & "', '" & jobname & "', '" & floorname & "', '" & tagname & "', '" & ShipDate & "', '" & Shiptime & "', '" & truck & "', '" & DescriptionName & "', '" & WindowValue & "' )"				
							'SQL = "INSERT INTO X_SHIP (BARCODE, JOB, FLOOR, TAG, SHIPDATE, ShipTime, TRUCK, Description) VALUES ('" & barcode & "', '" & jobname & "', '" & floorname & "', '" & tagname & "', '" & Date() & "', '" & Shiptime & "', '" & truck & "', '" & DescriptionName & "' )"				
							Scanned = True

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
							'------------Completed Entry
						end if
					end if
			end if
		end if
	end if
	rs.filter = ""
else 
		Scanned = FALSE
		errorDesc = "No Truck Entered for Scanning"

end if


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

	DbCloseAll

End Function

%>
   
   <!--New Form to collect the Job and Floor fields-->
   
   <%
   if (Len(barcode) = 0 AND Len(truck)= 0) then
		PanelClass = "panel"
   else 
		PanelClass = "panel3"
   end if
   %>
   
    <form id="AddTruck" title="Dock: <%response.write dockNum%>" class="<%response.write PanelClass%>" name="AddTruck" action="ShipTruckScan.asp" method="GET" selected="true">
        
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
	strSQL = FixSQLCheck("SELECT top 10 * FROM X_SHIP_TRUCK WHERE active = True ORDER BY ID DESC", isSQLServer)
	'strSQL = FixSQLCheck("SELECT top 1000 * FROM X_SHIP_TRUCK WHERE active = True {0} ORDER BY ID DESC", isSQLServer)
	
'	if truck <> "" and truck <> "-" then
'		strSQL = Replace(strSQL,"{0}"," AND ID=" & truck & " ",1)
'	else
'		strSQL = Replace(strSQL,"{0}","",1)
'	end if

	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	

'if truck <> "" and truck <> "-" then
'	rs.filter = "ID = " & truck
'	activeTruck = ""
'	if rs("truckName") <> "" then
'		activeTruck = rs("truckName") & " - "
'	end if
'	activeTruck = activeTruck & rs("sList")
'	response.write " <option value = '" & rs("id") & "'>" & activeTruck & "</option>"

'end if	
	if truck <> "" and truck <> "-" then
		rs.filter = "ID <> " & truck	
	else
		rs.filter = ""	
	end if
	rs.filter = "DockNum = " & dockNum
	do while not rs.eof
	Response.Write "<option value = '"
	Response.Write rs("id")
	Response.Write "'>"

	if rs("truckName") <> "" then
		Response.Write rs("truckName") & " - "
	end if
	Response.Write rs("sList") 
	Response.Write "</option>"
	rs.movenext
	loop

%>
</select>
	
	    </div>
		
		<div class="row">
                <label>Shipping Item</label>
                <input type="text" name='barcode' id='barcode' readonly>
				
        </div>
				<input type="hidden" name='dockNum' id='dockNum' value='<%response.write dockNum%>'>
        </fieldset>
        <BR>
		
            <ul id="Profiles" title="Active Trucks" selected="true">
        <%
		if Scanned = True then
			response.write" <li>Scan Successful</li>"
				if truck <> "" and truck <> "-" then
					rs.filter = "ID = " & truck
					response.write" <li>"  & Barcode & " On Truck: " & truck & "</li>"
					
				%>
				<script type="text/javascript">
				function FlashingScreenGreen() {
					
						beep = "Error"
						var flash = false;
						var task = setInterval(function() {
						if(flash = !flash) {
							document.body.style.backgroundColor = '#008000';
						} else {
							document.body.style.backgroundColor = '#00FF00';
						}
						}, 1000);
						
					}
					
					FlashingScreenGreen()
					
			  </script>
			  <%	

					rs.filter = ""
				else
					Response.write "<li> Please Select a Truck for Scanning </li>"
				end if
		else
			if barcode = "" then
			response.write" <li>No Item entered</li>"
			else
			if errorDesc = "Item Already Scanned" then
				%>
				<script type="text/javascript">
				function FlashingScreen() {
					
						beep = "Error"
						var flash = false;
						var task = setInterval(function() {
						if(flash = !flash) {
							document.body.style.backgroundColor = '#000';
						} else {
							document.body.style.backgroundColor = '#FFA500';
						}
						}, 1000);
						
					}
					
					FlashingScreen()
					
				 </script>
				  <%
			
			else
				%>
					<script type="text/javascript">
					function FlashingScreen() {
						
							beep = "Error"
							var flash = false;
							var task = setInterval(function() {
							if(flash = !flash) {
								document.body.style.backgroundColor = '#000';
							} else {
								document.body.style.backgroundColor = '#f00';
							}
							}, 1000);
							
						}
						
						FlashingScreen()
						
				  </script>
				  <%
				end if
				response.write" <li>" & errorDesc & "</li>"
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


'DbCloseAll

rs.close
set rs = nothing
'rs2.close
'set rs2 = nothing
DBConnection.close
set DBConnection = nothing

%>	

</body>
</html>