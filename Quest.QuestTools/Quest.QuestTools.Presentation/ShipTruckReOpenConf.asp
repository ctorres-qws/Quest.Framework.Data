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
  
  
</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Ship</a>
        </div>
   
<%
error = ""

Passkey = "SH1P"
Password = UCASE(TRIM(Request.Form("pwd")))

truck = Trim(Request.querystring("truck"))
dockNum = Trim(Request.querystring("dockNum"))
IsError = False

			truckNum = ""
			truckName = ""
			sList = ""

If (UCASE(Password) = Passkey) then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			If gstr_ErrMsg="" Then Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select
	
Else
%>

<form id="adminpass" title="Enter Password" class="panel" name="enter" action="ShipTruckReOpenConf.asp?Truck=<%response.write truck%>&dockNum=<%response.write dockNum%>" method="post" target="_self" selected="True">
<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
End If

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

if dockNum = "" then
	IsError = True 
	error = "A Dock must entered to Re-Open, Please retry"
else

	'RecordSet of all trucks
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_SHIP_TRUCK ORDER BY ID DESC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

	rs.filter = "ID = " & truck & " AND ACTIVE <> 0"
	if not rs.eof then
	else
	rs.filter = ""

		'First Check to see if Dock is Occupied	
		rs.filter = "dockNum = " & dockNum & " AND ACTIVE <> 0"
				if not rs.eof then
					IsError = true
					error = "Dock " & dockNum & " Already occupied with an Active truck"
				end if
			
		rs.filter = ""	
		
		if IsError = False then

		'Re-Open Closed Truck  ( Remove shipDate, set active = True, Set New Dock)

				'Set shipping Clear Statement
			SQL = FixSQLCheck("UPDATE X_SHIP_TRUCK SET shipDate = NULL , dockNum = '" & dockNum & "', TruckNum = TruckNum +1, active= TRUE  WHERE id = " & truck,isSQLServer)
				'Get a Record Set
			Set RS1 = DBConnection.Execute(SQL)
		
			' Un-Ship Items (Remove Shipdate from items)
		
				'Set shipDate Clear Statement
			'SQL2 = "UPDATE X_SHIP SET shipDate = NULL  WHERE truck = " & truck
				'Get a Record Set
			'Set RS2 = DBConnection.Execute(SQL2)

			'Details of RE-Opened Truck
			rs.filter = "ID = " & truck
			truckNum = rs("truckNum")
			truckName = rs("truckName")
			sList = rs("sList")
		end if
	end if
	
end if

DbCloseAll

End Function
 If (UCASE(Password) = Passkey) then
%>

  
   <!--New Form to collect the Job and Floor fields-->
	
            <ul id="Profiles" title="Re-Opened Truck" selected="true">
<%
if IsError = true then
	response.write "<li>Truck not Re-Opened due to Error:</li>"
	response.write "<li>" & error & "</li>"
	response.write " <a class='whiteButton' href=' ShipTruckReOpen.asp' target='_self'>Back to Re-Open Truck Form</a>"
else
	response.write "<li> Truck Reset to Active </li>"
	response.write "<li>Truck Name: " & truckName & "</li>"
	response.write "<li>Job/Floor: " & sList & "</li>"
	response.write "<li>Truck ID: " & truck & "</li>"
	response.write "<li>Re-open #: " & TruckNum & "</li>"

	response.write "<li>Truck Added on Dock: " & dockNum & "</li>"
end if


%>
			</ul>

<%
END IF
'rs.close
'set rs = nothing
'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>