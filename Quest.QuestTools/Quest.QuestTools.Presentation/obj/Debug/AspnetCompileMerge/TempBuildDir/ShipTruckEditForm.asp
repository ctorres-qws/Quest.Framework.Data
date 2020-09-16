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
		strSQL = "SELECT * FROM X_SHIP_TRUCK ORDER BY dockNum ASC"
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
        <a class="button leftButton" type="cancel" href="ShipTruckEdit.asp" target="_self">Select Truck</a>
    </div>
	
	    <form id="EditTruck" title="Update Shipping Trucks" class="panel" name="EditTruck" action="ShipTruckEditConf.asp" method="GET" target="_self" selected="true">
              
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
			
		<div class="row">    
			<label>JOB</label>
			<select name='sJob' id='sJob'  class ='leftinput'>
<%

				Set rsJob = Server.CreateObject("adodb.recordset")
				strSQL = "SELECT JOB FROM Z_JOBS Where COMPLETED = FALSE ORDER BY JOB ASC"
				rsJob.Cursortype = 2
				rsJob.Locktype = 3
				rsJob.Open strSQL, DBConnection

				rsJob.movefirst
				Do While Not rsJob.eof

				Response.Write "<option value='"
				Response.Write rsJob("JOB")
				Response.Write "'>"
				Response.Write rsJob("JOB")
				Response.write "</option>"

				rsJob.movenext

				loop
				rsJob.close
				set rsJob=nothing
%>
			</select>
		</div>
		
		<div class="row">
			<label>Floor</label>
			<input type="text" name='sFloor' id='sFloor' value='00'>
		</div>
	
		<div class="row">
		<table>
			<tr><td><button type='button' class="whiteButton" onclick="addList()">Add to Floor List</Button></td>
			<td><button type='button' class="whiteButton" onclick="removeList()">Remove from Floor List</Button></td>
			<td><button type='button' class="whiteButton" onclick="clearList()">Clear Floor List</Button></td>
			</tr>
		</table>
			<script>
			function addList() {
				var J = document.getElementById("sJob").value;
				var F = document.getElementById("sFloor").value;
				var CommaSearch = F.includes(",");
				if (CommaSearch == false ) {
					var JF = J.concat(F);
					var List = document.getElementById("sList").value;
					var JobList = List.split(',');
					var JobLimit = JobList.Length;
					var checker = "False";
					for (i = 0; i < JobList.length; i++) { 
						if (JobList[i] == JF) {
							checker = "True";
							i = 100;
						};
					};
					if (checker == "False") {
						if (List.length == 0 ) {
							document.getElementById("sList").value = List.concat(J,F);
						} else {
							document.getElementById("sList").value = List.concat(",",J,F,);
						};
					};
					btn = document.getElementById("Sbmt"); btn.disabled = false;
								
				} else {
					alert("Please Enter each Floor Individually, No Commas allowed");
				};
			}
			</script>
			
			<script>
			function removeList() {
				var J = document.getElementById("sJob").value;
				var F = document.getElementById("sFloor").value;
				var JF = J.concat(F);
				var List = document.getElementById("sList").value;
				var r = confirm("Remove " + JF + " from Acceptable List");
				if (r == true) {
					var JobList = List.split(',');
					var JobLimit = JobList.Length;
					var checker = "0";
					var NewList = ""
					for (i = 0; i < JobList.length; i++) { 
						if (JobList[i] == JF) {
							checker = i+1;
							i = 100;
						};
					};
					if (checker > 0) {
						for (i = 0; i < JobList.length; i++) { 
							if (i+1 == checker) {
							} else {
								if (NewList.length == 0 ) {
									NewList = NewList.concat(JobList[i]);
								} else {
									NewList = NewList.concat(",",JobList[i]);
								};
							};
						};
						document.getElementById("sList").value = NewList;
					};
					if (List2.length == 0 ) {
						btn = document.getElementById("Sbmt"); btn.disabled = true;
					};
				};
			document.getElementById("sList").value = List2
			}
			</script>
			
			<script>
			function clearList() {

				var r = confirm("Clear the Floors List!");
				if (r == true) {
					document.getElementById("sList").value = "" ;
					btn = document.getElementById("Sbmt"); btn.disabled = true;
					clearList.preventDefault();
				};

			}
			</script>
		</div>
	
		<div class="row">         
            <label>Floor List</label>
            <input type="text" name='sList' id='sList' value = '<% response.write rs("sList")%>' readonly>
		</div>
            					
			<div class="row" >
				<label>Require Ship Date</label>
				<%
				if RS("RequireDate") = "" then
					DateEnter = ""
				else
					YearNow = Year(rs("RequireDate"))
					MonthNow = Month(rs("RequireDate"))
					if Len(MonthNow) = 1 then
						MonthNow = "0" & MonthNow
					end if
					DayNow = Day(rs("RequireDate"))
					if Len(DayNow) = 1 then
						DayNow = "0" & DayNow
					end if
					DateEnter = YearNow & "-" & MonthNow & "-" & DayNow
				end if
				%>
				
				<input type="date" name='RequireDate' id='RequireDate' value = '<% response.write DateEnter%>'  class ='Long'>
			</div>	
			
						
			 </fieldset>
			 <h2>The following docks are Unavailable: <b><% Response.Write Docklimit %></b></h2>
			  <fieldset>
			<div class="row" >
				<label>Dock #</label>
				<select name='dockNum' id='dockNum'  class ='leftinput'>
				<option value=1 <%if rs("dockNum") = 1 then response.write "selected"%> >1</option>
				<option value=7 <%if rs("dockNum") = 7 then response.write "selected"%>>7</option>
				<option value=8 <%if rs("dockNum") = 8 then response.write "selected"%>>8</option>
				<option value=9 <%if rs("dockNum") = 9 then response.write "selected"%>>9</option>
				<option value=10 <%if rs("dockNum") = 10 then response.write "selected"%>>10</option>
				<option value=11 <%if rs("dockNum") = 11 then response.write "selected"%>>11</option>
				<option value=12 <%if rs("dockNum") = 12 then response.write "selected"%>>12</option>
				</select>
				
			</div>	
			
				<input type="hidden" name='tid' id='tid' value = '<% response.write tid%>' />
			
			<button class="whiteButton" id ="Sbmt" onClick="EditTruck.submit()">Update Truck</button><BR>
            
         
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
