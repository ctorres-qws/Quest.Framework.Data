<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Set a New Truck to Active-->
<!-- From Job/Floor create New or Next truck (truckNum = 1 or add the next truck) -->
<!-- Confirm by Dock that no other truck is not yet closed-->
<!-- Sets Truck to Active -->
<!-- Collect Data from ShippingTruckOpen.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Activate Truck</title>
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
   
<%
error = ""
truckName = Trim(Request.querystring("truckName"))
job = UCASE(Trim(Request.querystring("job")))
floor = UCASE(Trim(Request.querystring("floor")))
dockNum = Trim(Request.querystring("dockNum"))
IsError = False
currentDate = Date

Dim truckID, truckNum

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

DBOpen DBConnection, isSQLServer

if job = "" or floor = "" or dockNum = "" then
IsError = True 
error = "Job, Floor, and Dock must all be filled in to add a new truck, Please retry"
else

	'RecordSet of all trucks
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_Shipping_Truck_Test ORDER BY truckNum DESC"
		rs.Cursortype = GetDBCursorTypeInsert
		rs.Locktype = GetDBLockTypeInsert
		rs.Open strSQL, DBConnection

	'First Check to see if Dock is Occupied	
	rs.filter = " active = '1' AND dockNum = " & dockNum 
			if not rs.eof then
				IsError = true
				error = "Dock " & dockNum & " Already occupied with an Active truck"
			end if

	rs.filter = ""
	'Second Check to find truckNum
	rs.filter= "job= '" & job & "' AND floor = '" & floor & "'"
		if not rs.eof then
			rs.movefirst
			truckNum = rs("truckNum") + 1
		else 
			truckNum = 1
		end if
	if rs.eof then 
	truckID = 1
	else
		rs.movelast
		truckID = rs("ID") + 1
	end if

end if

If IsError = False Then
	'Set Truck Add Statement

	If gi_Mode = c_MODE_HYBRID Then
		If GetID(isSQLServer,2) <> "" Then 
			truckNum = GetID(isSQLServer,2)
		End If
	End If

	rs.AddNew
	rs.fields("job") = job
	rs.fields("floor") = floor
	rs.fields("truckNum") = truckNum
	rs.fields("dockNum") = dockNum
	rs.fields("active") = FixSQLBool("True", isSQLServer)
	rs.fields("truckName") = truckName
	rs.fields("CreateDate") = FixSQLDate(currentDate, isSQLServer)
	If GetID(isSQLServer,1) <> "" Then 
		rs.Fields("ID") = GetID(isSQLServer,1)
		truckID = GetID(isSQLServer,1)
	End if
	rs.Update

	Call StoreID1(isSQLServer, rs.Fields("ID"))
	Call StoreID2(isSQLServer, truckNum)

End If

DbCloseAll

End Function

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ShippingHomeTest.HTML" target="_self">Scan</a>
        </div>
<!--New Form to collect the Job and Floor fields-->
	<ul id="Profiles" title="Active Trucks" selected="true">
<%
if IsError = true then
	response.write "<li>Truck not added due to Error:</li>"
	response.write "<li>" & error & "</li>"
else
	response.write "<li> Truck Set to Active </li>"
	response.write "<li>Truck Name: " & truckName & "</li>"
	response.write "<li>Job: " & job & "</li>"
	response.write "<li>Floor: " & floor & "</li>"
	response.write "<li>Truck number for this Job/Floor: " & truckNum & "</li>"
	response.write "<li>Truck Added on Dock: " & dockNum & "</li>"
end if
response.write " <a class='whiteButton' href=' ShippingTruckOpenTest.asp' target='_self'>Back to Open Truck Form</a>"
response.write " <a class='whiteButton' href=' ShippingTruckScanTest.asp?truck=" & truckID & "&barcode=' target='_self'>Scan Items to Truck</a>"
response.write " <a class='whiteButton' href=' index.html#_Scan' target='_self'>Scan Menu</a>"

%>
	</ul>
<%
'rs.close
'set rs = nothing
'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>