<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- Finds the Job and Floor Requirements and matches it to the SHipping -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping View</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	    


    <%
jobname = Request.Querystring("Job")
fl = Request.QueryString("Floor")

flLength = Len(cstr(fl))
mode = request.QueryString("mode")
supplier = request.QueryString("supplier")
	Const adSchemaTables = 20
	
			Set rsShip = Server.CreateObject("adodb.recordset")
			strSQL = "SELECT * from X_SHIPPING"
			'rsShip.Cursortype = GetDBCursorType
			'rsShip.Locktype = GetDBLockType
			'rsShip.Open strSQL, DBConnection
			Set rsShip = GetDisconnectedRS(strSQL, DBConnection)

			Set rsTruck = Server.CreateObject("adodb.recordset")
			strSQL = "SELECT * from X_SHIPPING_TRUCK"
			'rsTruck.Cursortype = GetDBCursorType
			'rsTruck.Locktype = GetDBLockType
			'rsTruck.Open strSQL, DBConnection
			Set rsTruck = GetDisconnectedRS(strSQL, DBConnection)

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_ShipReport" target="_self">Ship View</a>
        </div>
   <ul id="Profiles" title="Glass Report - Commercial" selected="true">
        <li>Shipping Closed Trucks</li>
 <%
	response.write "<li><table border=1>"
    response.write "<tr><th>Job</th><th>Floor</th><th>Tag</th><th>Cut Status</th><th>Cut Date</th><th>Truck Name</th><th>Truck Code</th><th>Truck Ship Date</th></tr>"
	response.write "</tr>"

	Set rsTables = DBConnection.OpenSchema(adSchemaTables)
	Do While Not rsTables.eof
		If Left(CStr(rsTables("TABLE_NAME")), 4+3+flLength) = "Cut_" & jobname & fl then
		
			Set rsInsert = Server.CreateObject("adodb.recordset")
			strSQL = "SELECT DISTINCT [Line Item], cstatus, cdate FROM " & rsTables("TABLE_NAME") 
			
			rsInsert.Cursortype = GetDBCursorType
			rsInsert.Locktype = GetDBLockType
			rsInsert.Open strSQL, DBConnection
			
			Do while not rsInsert.eof
			LastTag = NewTag
			NewTag = rsInsert.fields("Line Item")
			if NewTag = LastTag then
			else
				response.write "<tr><td>" & Jobname & "</td>"
				response.write "<td>" & Fl & "</td>"
				response.write "<td>" & rsInsert.fields("Line Item") & "</td>"
				if rsInsert.fields("cstatus") = "0" then
					response.write "<td> Not Yet Cut </td>"
					response.write "<td> - </td>"
				else
					response.write "<td> Cut </td>"
					response.write "<td>" & rsInsert.fields("cdate") & "</td>"
				end if
				
				' Code for Shipping information
				' Find Shipping items
				rsShip.Filter = "JOB = '" & jobname & "' AND FLOOR = '" & fl & "' AND TAG = '" & rsInsert.fields("Line Item") & "'"
				if rsShip.eof then
					response.write "<td></td><td></td><td></td>"
				else
					'Find the Shipping Truck to match
					rsTruck.Filter = "ID = " & rsShip("Truck")
					if rsTruck.eof then
						response.write "<td></td><td></td><td></td>"
					else
						response.write "<td>" & rsTruck("TruckName") & "</td>"
						response.write "<td>" & rsTruck("job") & rsTruck("floor") & " - " & rsTruck("dockNum") & "</td>"
						response.write "<td>" & rsTruck("Shipdate") & "</td>"
					end if
				end if
			end if
			response.write "</tr>"
			rsInsert.movenext
			loop
			rsInsert.close
			set rsInsert = nothing
	   end if
	rsTables.movenext
	loop
rsTables.close
set rsTables = nothing

response.write "</table></li>"
response.write "<li> All Shipping items in: " & jobname & ", Floor: " & fl 
%>
<li><table>
<TR><TH>ITEMS</TH></tr>
<%

rsShip.Filter = "JOB = '" & jobname & "' AND FLOOR = '" & fl & "'"
do while not rsShip.eof
response.write "<tr><td>" & rsShip("Barcode") & "</td></tr>"
rsShip.movenext
loop
response.write "</table></li>"


rsShip.close
set rsShip = nothing
rsTruck.close
set rsTruck = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
       
</body>
</html>
