<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Close an Active truck-->
<!-- Sets Active to False and adds a Shipdate to the truck and to the items-->
<!-- Also Sets Backorder to true -->
<!-- Receives from  ShippingTruckClose.asp-->

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
   
<%
truck = Trim(Request.querystring("truck"))
currentDate = Date()
'currentDate = Now() '(Now includes Time)

'Close Truck  ( Add shipDate, set active = false)

	'Set Truck Add Statement

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

	SQL = FixSQLCheck("UPDATE X_SHIPPING_TRUCK SET shipDate = " & FormatDateToSQLCheck(currentDate,gstr_DateFormat,isSQLServer, "'") & ", active= FALSE  WHERE id = " & truck, isSQLServer)
		'Get a Record Set
	Set RS1 = DBConnection.Execute(SQL)

' Ship Items (Add Shipdate to items

	'Set Truck Add Statement
	SQL2 = "UPDATE X_SHIPPING SET shipDate = " & FormatDateToSQLCheck(currentDate,gstr_DateFormat,isSQLServer,"'") & "  WHERE truck = " & truck
	'Get a Record Set
	'Set RS2 = DBConnection.Execute(SQL2)

'BackOrder	set to true
backorderNum = request.querystring("back").count
dim Backorder()
ReDim Backorder(backorderNum)

For i=0 to backorderNum-1
	Backorder(i) = Request.Querystring("back")(i+1)
	
	'Set Backorder to True Statement
	SQL3 = FixSQLCheck("UPDATE X_SHIPPING SET BackOrder = True  WHERE id = " & Backorder(i),isSQLServer)

		'Get a Record Set
	Set RS3 = DBConnection.Execute(SQL3)
	
Next

	DbCloseAll

End Function

	Set DBConnection = Server.CreateObject("adodb.connection")
	DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp
	DBConnection.Open DSN

	'Closed Truck
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM X_Shipping_Truck WHERE ID = " & truck & " ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	'All Shipping Items
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_Shipping ORDER BY ID ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL2, DBConnection
%>
</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
       <a class="button leftButton" type="cancel" href="ShippingHome.HTML" target="_self">Scan</a>
        </div>

            <ul id="Profiles" title="Closed Truck" selected="true">
<%
If Not rs.EOF Then
	response.write "<li>Truck # " & rs("truckNum") & " for Job: " & rs("job") & " and Floor: " & rs("floor") & "</li>"
	TruckName = rs("job") & rs("floor")
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
			</ul>
			
			<!--#include file="ShippingTruckCloseGmail.asp"-->
<%
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
set DBConnection = nothing
%>	

</body>
</html>