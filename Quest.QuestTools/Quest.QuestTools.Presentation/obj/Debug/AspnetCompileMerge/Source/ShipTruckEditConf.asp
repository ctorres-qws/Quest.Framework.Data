<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Edit List to Manage ALL trucks on the X_SHIPPING_Truck table-->
<!-- X_SHIPPING_LIBRARY, X_SHIPPING_TRUCK, and X_Shipping Tables created at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Truck Maintainance allows Changing Docks, but must confirm Dock is open -->
<!-- Allows Truck details (not Number) to be edited - Dock, Job, Floor, Name
<!-- Inputs fromShippingTruckEditForm.asp-->
<!-- Inputs to ShippingTruckEditConf.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Truck</title>
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
tid = Trim(Request.Querystring("tid"))
truckName = Trim(Request.querystring("truckName"))
sList = UCASE(Trim(Request.querystring("sList")))
dockNum = Trim(Request.querystring("dockNum"))
RequireDate = Trim(Request.querystring("RequireDate"))
IsError = False
ReturnFloor = ""
ReturnFloorCode = ""

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

if sList = "" or dockNum = "" then
IsError = True 
error = "Job, Floor, and Dock must all be filled in to edit a truck, Please retry"
else

	'RecordSet of all trucks
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_SHIP_TRUCK ORDER BY DockNum DESC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

	
	'First Check to see if Dock is Occupied	
	rs.filter = " active = '1' AND ID <> " & tid & " AND dockNum = " & dockNum 
			if not rs.eof then
				IsError = true
				error = "Dock " & dockNum & " Already occupied with an Active truck"
			end if
		
	rs.filter = "ID = " & tid	
	if not rs.eof then
		'Second Check to see FloorList of the truck includes Already Scanned Items
		OldList = rs("sList")
		JobsList = Split(OldList, ",")
		NewList = sList
		Dim iJob(25)
		Dim iFloor(25)
		JobLimit = Ubound(JobsList)
		
		if (JobLimit => 1) Then 
			for i=0 to Ubound(JobsList)
				SplitItem = Trim(Jobslist(i))
				iJob(i) = Left(SplitItem,3)
				iFloor(i) = Right(SplitItem,(Len(SplitItem)-3))
				test = "7"
			next
		else
			if sList ="" then 
				JobLimit = -1
			else
				JobLimit = 0
				SplitItem = sList
				iJob(0) = Left(SplitItem,3)
				iFloor(0) = Right(SplitItem,(Len(SplitItem)-3))
			end if 
		end if
		
		
		for i=0 to Ubound(JobsList)
			JobFloor = iJob(i) & iFloor(i)
			if (instr(NewList,JobFloor) = 0) Then
				
			Set rsScan = Server.CreateObject("adodb.recordset")
			strSQLScan = "SELECT * FROM X_SHIP Where [Deleted] = 0 AND Truck =" & tid &" ORDER BY ID DESC"
			rsScan.Cursortype = 2
			rsScan.Locktype = 3
			rsScan.Open strSQLScan, DBConnection
			rsScan.Filter = "JOB = '" & iJob(i)& "' AND Floor = '" & iFloor(i) & "'"
			
			if not rsScan.eof then
				IsError = True 
				error = "Could Not Remove: " & iJob(i) & iFloor(i) & ". Windows are loaded onto this truck, Please Remove Windows to clear FloorList"
				
			end if 
			
			rsScan.close
			Set rsScan = nothing
			
			end if
		next
		
		
		rs.close
		set rs = nothing
	end if
	
	
	
	
end if
if IsError = False then
		'Set Truck Add Statement
	SQL2 = "UPDATE X_SHIP_TRUCK SET sList = '" & sList & "', dockNum = '" & dockNum & "', truckName= '" & truckName & "' WHERE ID = " & tid
		'Get a Record Set
	Set RS2 = DBConnection.Execute(SQL2)
	
	if isDate(RequireDate) then 
		SQL3 = "UPDATE X_SHIP_TRUCK SET RequireDate = " & FormatDateToSQLCheck(RequireDate,"YYYY-MM-DD",isSQLServer,"'") & " WHERE ID = " & tid
		'Get a Record Set
		Set RS3 = DBConnection.Execute(SQL3)
	end if
end if

DbCloseAll

End Function

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
	
            <ul id="Profiles" title="Active Trucks" selected="true">
<%
if IsError = true then
	response.write "<li>Truck not added due to Error:</li>"
	response.write "<li>" & error & "</li>"
else
	response.write "<li> Truck Updated </li>"
	response.write "<li>Truck Name: " & truckName & "</li>"
	response.write "<li>FloorsList: " & sList & "</li>"
	response.write "<li>Required Ship Date: " & RequireDate & "</li>"
	response.write "<li>Truck number for this Job/Floor: " & truckNum & "</li>"
	response.write "<li>Truck Updated on Dock: " & dockNum & "</li>"
end if	
response.write " <a class='whiteButton' href=' ShipTruckEdit.asp' target='_self'>Back to Manage Truck Menu</a>"

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