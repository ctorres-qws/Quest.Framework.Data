<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Set a New Truck to Active-->
<!-- From Job/Floor create New or Next truck (truckNum = 1 or add the next truck) -->
<!-- Confirm by Dock that no other truck is not yet closed-->
<!-- Sets Truck to Active -->
<!-- Confirms to ShippingTruckOpenConf.asp-->
<!--Date: September 10, 2019
	Modified By: Michelle Dungo
	Changes: Modified to change button type to <a> to stop page submitting twice
-->
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
   

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="AddTruck" title="Open Truck" class="panel" name="AddTruck" action="ShipTruckOpenConf.asp" method="GET" selected="true">
        
        <h2>Enter Truck Name, Then add Job and Floor Combinations to the Accepted List.</h2>
       <fieldset>

	    <div class="row">
                <label>Truck Name</label>
                <input type="text" name='truckName' id='truckName' >
        </div>
		
		<div class="row">    
			<label>JOB</label>
			<select name='sJob' id='sJob' class ='leftinput'">
<%

				Set rs = Server.CreateObject("adodb.recordset")
				strSQL = "SELECT JOB FROM Z_JOBS WHERE COMPLETED = FALSE ORDER BY JOB ASC"
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
%>
			</select>
		</div>
		
			
			
		<div class="row">
			<label>Floor</label>
			<input type="text" name='sFloor' id='sFloor' value='00'>
		</div>
		

		
		<div class="row">
		<table>
		<tr>
		<td><button  type='button' class="whiteButton" onclick="addList()">Add to Floor List</Button></td>
		<td><button  type='button' class="whiteButton" onclick="removeList()">Remove from Floor List</Button></td>
		<td>
		<button  type='button' class="whiteButton" onclick="clearList()">Clear Floor List</Button>
		</td>
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

				var r = confirm("Clear the Accepted List!");
				if (r == true) {
					document.getElementById("sList").value = "" ;
					btn = document.getElementById("Sbmt"); btn.disabled = true;
				};

			}
			</script>
		</div>
	
		<div class="row">         
            <label>Floor List</label>
            <input type="text" name='sList' id='sList' readonly>
		</div>
		
		<div class="row">
                <label>Required Shipping Date</label>			
                <input type="date" name='RequireDate' id='RequireDate' class="Long">
        </div>
		
		<div class="row">
			<label>Dock</label>
			
			<select name='dockNum' id='dockNum' class ='leftinput'>
				<option value=1>1</option>
				<option value=7>7</option>
				<option value=8>8</option>
				<option value=9>9</option>
				<option value=10>10</option>
				<option value=11>11</option>
				<option value=12>12</option>
			</select>
        </div>

		</fieldset>

        <BR>
		<a class="whiteButton" id="Sbmt" onClick="AddTruck.submit()" disabled >Submit</a><BR>
		
		
            <ul title="Active Trucks" selected="true">
        <%


	'View all the current Active Trucks and Docks (So new truck is not added twice)
	

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = FixSQL("SELECT * FROM X_SHIP_TRUCK WHERE active = True ORDER BY dockNum ASC")
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	
	
			response.write "<h3>Currently Active Trucks</h3>"
			response.write "<table border='1' class='sortable'><tr><th>Dock</th><th>Truck Name</th><th>Job/Floor</th><th>Truck #</th></tr>"
		
		if not rs.bof then
		rs.movefirst
		do while not rs.eof
			
			Response.write "<tr>"
			Response.write "<td>" & trim(RS("dockNum")) & "</td>"
			Response.write "<td>" & trim(RS("truckName")) & "</td>"
			Response.write "<td>" & trim(RS("sList")) & "</td>"
			Response.write "<td>" & trim(RS("ID")) & "</td>"
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