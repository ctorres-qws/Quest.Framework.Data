<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Set a New Truck to Active-->
<!-- From Job/Floor create New or Next truck (truckNum = 1 or add the next truck) -->
<!-- Confirm by Dock that no other truck is not yet closed-->
<!-- Sets Truck to Active -->
<!-- Confirms to ShippingTruckOpenConf.asp-->

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
   

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ShippingHome.HTML" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="AddTruck" title="Add Truck" class="panel" name="AddTruck" action="ShippingTruckOpenConf.asp" method="GET" selected="true">
        
        <h2>Input Job and Floor of the New Truck </h2>
       <fieldset>

	    <div class="row">
                <label>Truck Name</label>
                <input type="text" name='truckName' id='truckName' >
        </div>
	   
        <div class="row">
                   <label>JOB</Label>
<select name="Job">
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT JOB FROM Z_JOBS ORDER BY JOB ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.movefirst
Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("JOB")
Response.Write "'>"
Response.Write rs("JOB")
Response.write "</option>"

rs.movenext

loop
rs.close
set rs=nothing
%></select>
        </div>
		
		<div class="row">
                <label>Floor</label>
                <input type="text" name='floor' id='floor' >
        </div>
		</fieldset>
		<h2> Input Dock Number for the truck</h2>
		<fieldset>	
		<div class="row">
                <label>Dock</label>
                <input type="text" name='dockNum' id='dockNum' >
        </div>
 
        </fieldset>
        <BR>
		<a class="whiteButton" onClick="AddTruck.submit()">Submit</a><BR>
		
		
            <ul id="Profiles" title="Active Trucks" selected="true">
        <%


	'View all the current Active Trucks and Docks (So new truck is not added twice)
	

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = FixSQL("SELECT * FROM X_Shipping_Truck WHERE active = True ORDER BY dockNum ASC")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	
	
			response.write "<h3>Currently Active Trucks</h3>"
			response.write "<table border='1' class='sortable'><tr><th>Dock</th><th>Truck Name</th><th>Job</th><th>Floor</th><th>Truck #</th></tr>"
		
		if not rs.bof then
		rs.movefirst
		do while not rs.eof
			
			Response.write "<tr>"
			Response.write "<td>" & trim(RS("dockNum")) & "</td>"
			Response.write "<td>" & trim(RS("truckName")) & "</td>"
			Response.write "<td>" & trim(RS("Job")) & "</td>"
			Response.write "<td>" & trim(RS("Floor")) & "</td>"
			Response.write "<td>" & trim(RS("truckNum")) & "</td>"
			Response.write "</tr>"
		rs.movenext
		loop
		else
			response.write "<tr ><td colspan = '5'>No Active Trucks</td></tr>"
		end if
	response.write "</table>"
	rs.close
	set rs = nothing

    
        %>
			</ul>
            </form>
		

<%

DBConnection.close
set DBConnection = nothing
%>	

</body>
</html>