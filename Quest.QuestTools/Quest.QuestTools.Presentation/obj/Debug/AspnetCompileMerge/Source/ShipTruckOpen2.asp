<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 2014, by Michael Bernholtz - Set a New Truck to Active-->
<!-- From Job/Floor create New or Next truck (truckNum = 1 or add the next truck) -->
<!-- Confirm by Dock that no other truck is not yet closed-->
<!-- Sets Truck to Active -->
<!-- Confirms to ShipTruckOpenConf.asp-->
<!-- Michelle Dungo attempt to Create Job Floor - incomplete-->

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
  <style>
.csButton { margin-right: 30px; height: 30px !important; padding: 10px 10px;}  
  </style>
<SCRIPT language="javascript">
		function val() {
			var d = document.getElementById("select_id").value;
			//'alert(d);
			//$('#MasterID').val(d);
			//alert($('#MasterID').val());
		}
		
		function getFloor() {
		var floor=document.getElementById('floorList');
		
		<%

		'Set rs = Server.CreateObject("adodb.recordset")
		'strSQL = "SELECT Floor FROM " & d & " ORDER BY JOB ASC"
		'rs.Cursortype = 2
		'rs.Locktype = 3
		'rs.Open strSQL, DBConnection

		'opCounter = 0
		'rs.movefirst
		'Do While Not rs.eof
			'floor.options[opCounter]= new Option('
		'Response.Write "<option value='"
'Response.Write rs("FLOOR")
'Response.Write "'>"
'Response.Write rs("FLOOR")
'Response.write "</option>"

'rs.movenext

'loop
'rs.close
'set rs=nothing
%>
		}
		
		function addRow(tableID) {

			var table = document.getElementById(tableID);

			var rowCount = table.rows.length;
			var row = table.insertRow(rowCount);

			var colCount = table.rows[0].cells.length;

			for(var i=0; i<colCount; i++) {

				var newcell	= row.insertCell(i);

				newcell.innerHTML = table.rows[0].cells[i].innerHTML;
				//alert(newcell.childNodes);
				switch(newcell.childNodes[0].type) {
					case "text":
							newcell.childNodes[0].value = "";
							break;
					case "checkbox":
							newcell.childNodes[0].checked = false;
							break;
					case "select-one":
							newcell.childNodes[0].selectedIndex = 0;
							break;
				}
			}
		}

		function deleteRow(tableID) {
			try {
			var table = document.getElementById(tableID);
			var rowCount = table.rows.length;

			for(var i=0; i<rowCount; i++) {
				var row = table.rows[i];
				var chkbox = row.cells[0].childNodes[0];
				if(null != chkbox && true == chkbox.checked) {
					if(rowCount <= 1) {
						alert("Cannot delete all the rows.");
						break;
					}
					table.deleteRow(i);
					rowCount--;
					i--;
				}


			}
			}catch(e) {
				alert(e);
			}
		}

	</SCRIPT>
   

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ShippingHomeTest.HTML" target="_self">Scan</a>
        </div>
   
   <!--New Form to collect the Job and Floor fields-->
    <form id="AddTruck" title="Add Truck" class="panel" name="AddTruck" action="ShippingTruckOpenConfTest.asp" method="GET" selected="true">
        
        <h2>Input Job and Floor of the New Truck </h2>
       <fieldset>

	    <div class="row">
                <label>Truck Name</label>
                <input type="text" name='truckName' id='truckName' >
        </div>
		<div class="row">
		<table border="0" cellpadding="10">
			<tr>
				<td><label><strong>Select Job and Floor</strong></label></td>			
			</tr>	
			<tr>		
				<td><a class="csButton" type="button" value="Add Row" onclick="addRow('dataTable')">Add Row</a></td>		
				<td><a class="csButton" type="button" value="Delete Row" onclick="deleteRow('dataTable')">Delete Row</a></td>
			</tr>
		</table>
		</div>
        <div class="row">
	<TABLE id="dataTable" width="350px" border="0">
		<TR>
			<TD><INPUT type="checkbox" name="chk"/></TD>
			<TD>				   
<select name="Job" id="select_id" onchange="val()" onclick="getFloor()">
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
</TD>
			<TD>	
<input type="hidden" id="MasterID" name="MasterID" value="">			
<select name="Floor" id="floorList">
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Floor FROM " & "AAA" & " ORDER BY Floor ASC"
'strSQL = "SELECT Floor FROM " & d & " ORDER BY Floor ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.movefirst
Do While Not rs.eof
Response.Write "<option value='"
Response.Write rs("FLOOR")
Response.Write "'>"
Response.Write rs("FLOOR")
Response.write "</option>"

rs.movenext

loop
rs.close
set rs=nothing
%></select>
</TD>
		</TR>
	</TABLE>

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
	strSQL = FixSQL("SELECT * FROM X_SHIP_TRUCK WHERE active = True ORDER BY dockNum ASC")
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