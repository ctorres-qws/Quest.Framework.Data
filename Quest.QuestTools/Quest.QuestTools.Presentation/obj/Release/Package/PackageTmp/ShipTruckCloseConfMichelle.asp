<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Close an Active truck-->
<!-- Sets Active to False and adds a Shipdate to the truck and to the items-->
<!-- Also Sets Backorder to true -->
<!-- Receives from  ShipTruckClose.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Close Truck</title>
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
		<h1 id="pageTitle">Close Truck</h1>
		<a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Ship</a>
	</div>

        
   
<%
Passkey = "SH1P"
truck = Trim(Request.querystring("truck"))
Password = UCASE(TRIM(Request.Form("pwd")))
currentDate = Date()
'currentDate = Now() '(Now includes Time)

If (UCASE(Password) = Passkey) then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			'Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

Else
%>

<form id="adminpass" title="Enter Password" class="panel" name="enter" action="ShipTruckCloseConfMichelle.asp?Truck=<%response.write truck%>" method="post" target="_self" selected="True">
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

	SQL = FixSQLCheck("UPDATE X_SHIP_TRUCK_MICHELLE SET shipDate = " & FormatDateToSQLCheck(currentDate,gstr_DateFormat,isSQLServer, "'") & ", active= FALSE  WHERE id = " & truck, isSQLServer)
		'Get a Record Set
	Set RS1 = DBConnection.Execute(SQL)

' Ship Items (Add Shipdate to items

	'Set Truck Add Statement
	SQL2 = "UPDATE X_SHIP_MICHELLE SET shipDate = " & FormatDateToSQLCheck(currentDate,gstr_DateFormat,isSQLServer,"'") & "  WHERE truck = " & truck
	'Get a Record Set
	'Set RS2 = DBConnection.Execute(SQL2)

'BackOrder	set to true
'backorderNum = request.querystring("back").count
'dim Backorder()
'ReDim Backorder(backorderNum)

'For i=0 to backorderNum-1
'	Backorder(i) = Request.Querystring("back")(i+1)
'	
'	'Set Backorder to True Statement
'	SQL3 = FixSQLCheck("UPDATE X_SHIP_MICHELLE SET BackOrder = True  WHERE id = " & Backorder(i),isSQLServer)
'
'		'Get a Record Set
'	Set RS3 = DBConnection.Execute(SQL3)
'	
'Next

	DbCloseAll

End Function

	Set DBConnection = Server.CreateObject("adodb.connection")
	DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp
	DBConnection.Open DSN

	'Closed Truck
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM X_SHIP_TRUCK_MICHELLE WHERE ID = " & truck & " ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	'All Shipping Items
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_SHIP_MICHELLE WHERE [DELETED] = FALSE ORDER BY ID ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL2, DBConnection

If (UCASE(Password) = Passkey) then
%>
  <ul id="Profiles" title="Closed Truck" selected="TRUE" >
<%
	If Not rs.EOF Then
		response.write "<li>Truck for JobFloor: " & rs("sList") & "</li>"
		TruckName = rs("sList")
		If rs("truckName") <> "" Then
			response.write "<li>Named: " & rs("truckName") & "</li>"
		End If
		response.write "<li>Closed on:" & currentDate & "</li>"
		If BackorderNum >0 Then
			response.write "<br><li>Items added to Backorder:</li>"
			For i=0 to backorderNum-1
				rs2.filter = "id = " & Backorder(i)
				response.write "<li>" & rs2("job") & "-" & rs2("floor") & rs2("tag") & "</li> "
				rs2.filter = ""
			Next
		End If
	Else
		Response.Write("<br />&nbsp;Truck Not Found")
	End If


	%>
	<!--#include file="ShipTruckCloseGmailMichelle.asp"-->
				</ul>
				
				
	<%

End if
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs1.close
set rs1 = nothing
DBConnection.close
set DBConnection = nothing
%>	

</body>
</html>