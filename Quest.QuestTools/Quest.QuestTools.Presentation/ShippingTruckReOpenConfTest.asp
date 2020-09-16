<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created at Request of Socol, in the rare event a truck is closed and then called back  -->
<!-- Created April 2014, by Michael Bernholtz - ReOpen a Closed truck-->
<!-- Confirms correct Dock for Re-Open -->
<!-- Sets Active to True and removes Shipdate from the truck-->
<!-- Confirms to ShippingTruckReopenConf.asp-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Re-Open Truck</title>
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
truck = Trim(Request.querystring("truck"))
dockNum = Trim(Request.querystring("dockNum"))
IsError = False

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		If gstr_ErrMsg="" Then Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

if dockNum = "" then
	IsError = True 
	error = "A Dock must entered to Re-Open, Please retry"
else

	'RecordSet of all trucks
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_Shipping_Truck_Test ORDER BY truckNum DESC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

	
	'First Check to see if Dock is Occupied	
	rs.filter = " active = '1' AND dockNum = " & dockNum 
			if not rs.eof then
				IsError = true
				error = "Dock " & dockNum & " Already occupied with an Active truck"
			end if
		
	rs.filter = ""	
	
	if IsError = False then

	'Re-Open Closed Truck  ( Remove shipDate, set active = True, Set New Dock)

			'Set shipping Clear Statement
		SQL = FixSQLCheck("UPDATE X_SHIPPING_TRUCK_TEST SET shipDate = NULL , dockNum = '" & dockNum & "', active= TRUE  WHERE id = " & truck,isSQLServer)
			'Get a Record Set
		Set RS1 = DBConnection.Execute(SQL)
	
		' Un-Ship Items (Remove Shipdate from items)
	
			'Set shipDate Clear Statement
		SQL2 = "UPDATE X_SHIPPING_TEST SET shipDate = NULL  WHERE truck = " & truck
			'Get a Record Set
		Set RS2 = DBConnection.Execute(SQL2)

		'Details of RE-Opened Truck
		rs.filter = "ID = " & truck

		truckName = rs("truckName")
		job = rs("job")
		floor = rs("floor")
		truckNum = rs("truckNum")

	end if
end if

DbCloseAll

End Function

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ShippingHomeTest.HTML" target="_self">Ship</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
	
            <ul id="Profiles" title="Re-Opened Truck" selected="true">
<%
if IsError = true then
	response.write "<li>Truck not Re-Opened due to Error:</li>"
	response.write "<li>" & error & "</li>"
	response.write " <a class='whiteButton' href=' ShippingTruckReOpenTest.asp' target='_self'>Back to Re-Open Truck Form</a>"
else
	response.write "<li> Truck Reset to Active </li>"
	response.write "<li>Truck Name: " & truckName & "</li>"
	response.write "<li>Job: " & job & "</li>"
	response.write "<li>Floor: " & floor & "</li>"
	response.write "<li>Truck number for this Job/Floor: " & truckNum & "</li>"
	response.write "<li>Truck Added on Dock: " & dockNum & "</li>"
end if


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