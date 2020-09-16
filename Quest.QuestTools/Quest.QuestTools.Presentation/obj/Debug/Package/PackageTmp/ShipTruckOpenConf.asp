<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Set a New Truck to Active-->
<!-- From Job/Floor create New or Next truck (truckNum = 1 or add the next truck) -->
<!-- Confirm by Dock that no other truck is not yet closed-->
<!-- Sets Truck to Active -->
<!-- Collect Data from ShipTruckOpen.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Open New Truck</title>
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
TruckID = ""
truckName = Trim(Request.querystring("truckName"))
sList = UCASE(Trim(Request.querystring("sList")))
RequireDate = Request.querystring("RequireDate")
if RequireDate ="" then
	RequireDate = DATE
end if 
dockNum = Trim(Request.querystring("dockNum"))
IsError = False
currentDate = Date
EnterNum = 1


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
IsError = False
if sList = "" or dockNum = "" then
	IsError = True 
	error = "Job, Floor, and Dock must all be filled in to add a new truck, Please retry"
else

	'RecordSet of all trucks
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_SHIP_TRUCK ORDER BY ID DESC"
		rs.Cursortype = GetDBCursorTypeInsert
		rs.Locktype = GetDBLockTypeInsert
		rs.Open strSQL, DBConnection
	
	rs.filter = ""
	rs.filter = "TruckName = '" & truckName & "' AND sList = '" & sList & "'"
	
	if not rs.eof then
		IsError= True
		error = "Truck Already Added"
		TruckID = rs("ID")
	end if

	'First Check to see if Dock is Occupied	
	rs.filter = ""
	rs.filter = "dockNum = " & dockNum & " AND ACTIVE <> 0"

	if not rs.eof then
		IsError = True 
		error = "Dock " & dockNum & " Already occupied with an Active truck"
	end if
	
end if
rs.filter = ""
rs.filter = "TruckName = '" & truckName & "' AND sList = '" & sList & "'AND CreateDate = #" & FixSQLDate(currentDate, isSQLServer) & "#"
if rs.eof then

	If IsError = False Then
		'Set Truck Add Statement
		rs.AddNew
		rs.fields("sList") = sList
		rs.fields("dockNum") = dockNum
		rs.fields("active") = FixSQLBool("True", isSQLServer)
		rs.fields("truckName") = truckName
		rs.fields("truckNum") = 0
		rs.fields("RequireDate") = FixSQLDate(RequireDate, isSQLServer)
		rs.fields("CreateDate") = FixSQLDate(currentDate, isSQLServer)
		If GetID(isSQLServer,1) <> "" Then 
			rs.Fields("ID") = GetID(isSQLServer,1)
			truckID = GetID(isSQLServer,1)
		End if
		rs.Update
		EnterNum = EnterNum + 1

		rs.movelast
		Call StoreID1(isSQLServer, rs.Fields("ID"))
	End if
End If

DbCloseAll

End Function

If EnterNum >=2 then
	IsError = False
end if

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">Open Truck</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Scan</a>
        </div>
<!--New Form to collect the Job and Floor fields-->
	<ul id="Profiles" title="Active Trucks" selected="true">
<%
if IsError = True  then
	response.write "<li>Truck not added due to Error:</li>"
	response.write "<li>" & error  &  "</li>"
else
	response.write "<li> Truck Set to Active </li>"
	response.write "<li>Truck Name: " & truckName & "</li>"
	response.write "<li>Truck ID: " & truckID & "</li>"
	response.write "<li>List of Acceptable Locations: " & sList & "</li>"
	response.write "<li>Truck Added on Dock: " & dockNum & "</li>"
	response.write "<li>This truck should be Closed by: " & RequireDate & "</li>"
end if
response.write " <a class='whiteButton' href=' ShipTruckOpen.asp' target='_self'>Back to Open Truck Form</a>"
response.write " <a class='greenButton' href=' ShipTruckOpenConfPrint.asp?Truck=" & truckID &"' target='_self'>Print Truck Pages</a>"
response.write " <a class='whiteButton' href=' index.html#_Ship' target='_self'>Ship Menu</a>"

%>
	</ul>
</body>
</html>