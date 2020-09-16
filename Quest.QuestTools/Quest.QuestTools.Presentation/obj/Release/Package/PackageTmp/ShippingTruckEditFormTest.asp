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
tid = trim(UCASE(Request.Querystring("tid")))

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM X_SHIPPING_TRUCK_TEST ORDER BY dockNum ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		rs.filter = "active = TRUE"
		' Note for Adding Dock, to not add where another truck already is sitting
		if rs.eof then
			Docklimit = "All Docks Available"
		else
			rs.movefirst
			Docklimit = trim(rs("dockNum"))
			rs.movenext
			do while not rs.eof
				Docklimit = Docklimit & ", " & trim(rs("dockNum"))
				rs.movenext
			loop
		end if
		rs.filter = ""
		
		' Change filter to show the truck that was selected
		rs.filter = "ID = " & tid
		
%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Manage Truck</h1>
        <a class="button leftButton" type="cancel" href="ShippingTruckEditTest.asp" target="_self">Select Truck</a>
    </div>
	
	    <form id="EditTruck" title="Update Shipping Trucks" class="panel" name="EditTruck" action="ShippingTruckEditConfTest.asp" method="GET" target="_self" selected="true">
              
		<h2>Edit Truck Details</h2>
			  
        <fieldset>               

			<div class="row" >
				<label>Truck Name</label>
				<%
				if not isNull(rs("truckName")) then
					truckName = Replace(rs("truckName"), " ","&nbsp;")
				else
					truckName = ""
				end if
				%>
				<input type="Text" name='truckName' id='truckName' value= <% response.write truckName %> >
			</div>	
			
			<div class="row" >
				<label>Job</label>
				<input type="Text" name='Job' id='Job' value=<% response.write rs("job")%> >
			</div>
            		
			<div class="row" >
				<label>Floor</label>
				<input type="Text" name='Floor' id='Floor' value= <% response.write rs("floor")%>>
			</div>			
			
			<div class="row" >
				<label>Ship Date</label>
				<input type="text" name='ShipDate' id='ShipDate' value = <% response.write rs("ShipDate")%>>
			</div>	
			
						
			 </fieldset>
			 <h2>The following docks are Unavailable: <b><% Response.Write Docklimit %></b></h2>
			  <fieldset>
			<div class="row" >
				<label>Dock #</label>
				<input type="number" name='dockNum' id='dockNum' value = <% response.write rs("dockNum")%>>
			</div>	
			
				<input type="hidden" name='tid' id='tid' value = '<% response.write tid%>'>
			
			<a class="whiteButton" href="javascript:EditTruck.submit()">Update Truck</a>
            
         
		</fieldset>


            
    </form>

  </ul>          
     
<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>
               
</body>
</html>
