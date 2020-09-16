                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Truck View</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<meta http-equiv="refresh" content="1200" >
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	    


    <%
	''<!--#include file="todayandyesterday.asp"-->
	
	Truck = Request.QueryString("truck")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING WHERE [Truck] = " & Truck & " ORDER BY job, floor, TAG ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar"> 
 
        <ul id="Profiles" title="Window Report" selected="true">
	
	<li>Truck Closed: <%response.write TruckName%> - #<%response.write Truck%> - <%response.write Now%></li>
	
	
<% 
AllCounter = 0	
TruckNumbers = ""
do while not rs.eof
	LastJob = UCASE(CurrentJob)
	LastFloor = UCASE(CurrentFloor)
	CurrentJob = UCASE(rs("JOB"))
	CurrentFloor = UCASE(rs("FLOOR"))


	if (CurrentJob = LastJob) and (CurrentFloor = LastFloor) then
		Counter = Counter + 1
	else 
		if LastJob <> "" then
				Set rs2 = Server.CreateObject("adodb.recordset")
				strSQL2 = "SELECT * FROM " & LastJob & " WHERE [Floor] = '" & LastFloor & "' ORDER BY TAG ASC"
				rs2.Cursortype = 2
				rs2.Locktype = 3
				rs2.Open strSQL2, DBConnection

				TotalCounter = 0
				Do until rs2.eof
					TotalCounter = TotalCounter + 1
				rs2.movenext
				loop
				rs2.close
				set rs2 = nothing
				
				Set rs3 = Server.CreateObject("adodb.recordset")
				strSQL3 = "SELECT * FROM X_SHIPPING WHERE [Job] = '" & LastJob & "' AND [Floor] = '" & LastFloor & "' ORDER BY TRUCK, TAG ASC"
				rs3.Cursortype = 2
				rs3.Locktype = 3
				rs3.Open strSQL3, DBConnection
				TruckCounter = 0
				TruckNumbers = ""
				LastTruck = ""
				NewTruck = ""
				Do until rs3.eof
					LastTruck = NewTruck
					NewTruck = rs3("Truck")
					if NewTruck = LastTruck then
					else
							if TruckNumbers = "" then
							TruckNumbers = NewTruck
							else
							TruckNumbers = TruckNumbers & "; " & NewTruck
							end if
					End if
					TruckCounter = TruckCounter + 1
				rs3.movenext
				loop
				rs3.close
				set rs3 = nothing
				
				response.write "<li>" & LastJob & " " & LastFloor & " " & Counter & " / " & TotalCounter & " (" & TruckCounter & ") "
				response.write " - All trucks containing "& LastJob & " " & LastFloor & ": "& TruckNumbers & "</li>"
		end if
		Counter = 1
	end if
AllCounter = AllCounter + 1
rs.movenext
loop

			if LastJob <> "" then
				Set rs2 = Server.CreateObject("adodb.recordset")
				strSQL2 = "SELECT * FROM " & LastJob & " WHERE [Floor] = '" & LastFloor & "' ORDER BY TAG ASC"
				rs2.Cursortype = 2
				rs2.Locktype = 3
				rs2.Open strSQL2, DBConnection

				TotalCounter = 0
				Do until rs2.eof
					TotalCounter = TotalCounter + 1
				rs2.movenext
				loop
				rs2.close
				set rs2 = nothing
				
				Set rs3 = Server.CreateObject("adodb.recordset")
				strSQL3 = "SELECT * FROM X_SHIPPING WHERE [Job] = '" & LastJob & "' AND [Floor] = '" & LastFloor & "' ORDER BY TRUCK, TAG ASC"
				rs3.Cursortype = 2
				rs3.Locktype = 3
				rs3.Open strSQL3, DBConnection
				TruckCounter = 0
				TruckNumbers = ""
				LastTruck = ""
				NewTruck = ""
								
				Do until rs3.eof
					LastTruck = NewTruck
					NewTruck = rs3("Truck")
					if NewTruck = LastTruck then
					else
							if TruckNumbers = "" then
							TruckNumbers = NewTruck
							else
							TruckNumbers = TruckNumbers & "; " & NewTruck
							end if
					End if
					TruckCounter = TruckCounter + 1
				rs3.movenext
				loop
			
				rs3.close
				set rs3 = nothing
				
				response.write "<li>" & LastJob & " " & LastFloor & " " & Counter & " / " & TotalCounter & " (" & TruckCounter & ") "
				response.write " - All trucks containing "& LastJob & " " & LastFloor & ": "& TruckNumbers & "</li>"
			end if	
	
rs.movefirst
response.write "<li> Total Windows: " & AllCounter & "</li>"

response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>TRUCK</th></tr>"
do while not rs.eof

	response.write "<tr>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("TRUCK") & "</td>"
	response.write "</tr>"

	rs.movenext
loop
response.write "</table></li>"


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
  
</body>
</html>
