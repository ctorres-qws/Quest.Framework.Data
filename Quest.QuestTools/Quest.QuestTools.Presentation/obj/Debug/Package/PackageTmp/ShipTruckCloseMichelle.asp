<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Close an Active truck-->
<!-- Dropdown of active Trucks to select  -->
<!-- Sets Active to False and adds a Shipdate to the truck-->
<!-- Confirms to ShippingTruckCloseConf.asp-->

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
truck = Request.Querystring("Truck")
	'All Trucks (Active and Inactive)
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM X_SHIP_TRUCK_MICHELLE  ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

%>
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ShipHome.HTML" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="CloseTruck" title="Close Truck" class="panel" name="CloseTruck" action="ShipTruckCloseConfMichelle.asp" method="GET" selected="true">
        
        <h2>Select Active Truck for Closing</h2>
       <fieldset>

        <div class="row">
            <label>Truck</label>
            <select name="truck" id="truck">
<%

if truck <> "" and truck <> "0" then
	rs.filter = "ID = " & truck
	if rs("truckName") <> "" then
		activeTruck = rs("truckName") & " - "
	end if
	activeTruck = activeTruck & rs("sList") & "-" & rs("ID")
	response.write " <option value = '" & rs("id") & "'>" & activeTruck & "</option>"
else
	response.write " <option value = ''>Choose a Truck</option>"
end if	
	
rs.filter = ""	
rs.filter = "Active = TRUE"
rs.movefirst
do while not rs.eof
Response.Write "<option value = '"
Response.Write rs("id")
Response.Write "'>"
	if rs("truckName") <> "" then
		Response.write rs("truckName") & " - "
	end if
Response.Write rs("sList") 
rs.movenext
loop

%>
</select>
	
	    </div>
				
        </fieldset>
        <BR>
			<a class="redButton" onClick="CloseTruck.action='ShipTruckCloseConfMichelle.asp'; CloseTruck.submit()">Confirm and Close Truck</a><BR>
		   </form>

<%
rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>
</body>
</html>