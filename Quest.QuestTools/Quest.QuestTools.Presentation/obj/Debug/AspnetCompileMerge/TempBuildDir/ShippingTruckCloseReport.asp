                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->
<!-- Floors that close without all windows show in Red-->

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

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * FROM X_SHIPPING_TRUCK WHERE [ID] = " & Truck & " Order by ID "
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection
TruckName = rs4("TruckName")
ShipDate = rs4("ShipDate")

rs4.close
set rs4 = nothing

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar"> 
	
 
        <ul id="Profiles" title="Window Report" selected="true">
	<BR>
	<BR>
	<li>Truck Closed: <%response.write TruckName%> - #<%response.write Truck%> - <%response.write ShipDate%></li>
	
	
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
								TruckNumbers = "<a href='http://172.18.13.31:8081/ShippingTruckCloseReport.asp?Truck=" & NewTruck & "' target='_self'>T " & NewTruck & "</a>"
								'TruckNumbers = "T" & NewTruck
							else
								TruckNumbers = TruckNumbers & "; <a href='http://172.18.13.31:8081/ShippingTruckCloseReport.asp?Truck=" & NewTruck & "' target='_self'>T " & NewTruck & "</a>"
								'TruckNumbers = TruckNumbers & "; T" & NewTruck
							end if
					End if
					TruckCounter = TruckCounter + 1
				rs3.movenext
				loop
				rs3.close
				set rs3 = nothing
				
				
				
				response.write "<li><a href='http://172.18.13.31:8081\ShippingTruckFullReport.asp?Job=" & LastJob & "&Floor=" & LastFloor & "' target='_blank'>" & LastJob &  " " & LastFloor & "</a>" 
				
				
				
				
				if TruckCounter = TotalCounter then
					response.write " " & TruckCounter & " / " & TotalCounter & " (" & Counter & ") "
				else	
					response.write " <font color='red'><b>" & TruckCounter & " / " & TotalCounter & " (" & Counter & ") </b></font>"
				end if
				
				response.write " - All trucks containing "& LastJob & " " & LastFloor & ": "& TruckNumbers & "</li>"
		end if
		Counter = 1
	end if
if rs("Window") = "Window" then 
AllCounter = AllCounter + 1
end if
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
								TruckNumbers = "<a href='http://172.18.13.31:8081/ShippingTruckCloseReport.asp?Truck=" & NewTruck & "' target='_self'>T " & NewTruck & "</a>"
								'TruckNumbers = "T" & NewTruck
							else
								TruckNumbers = TruckNumbers & "; <a href='http://172.18.13.31:8081/ShippingTruckCloseReport.asp?Truck=" & NewTruck & "' target='_self'>T " & NewTruck & "</a>"
								'TruckNumbers = TruckNumbers & "; T" & NewTruck
							end if
					End if
					TruckCounter = TruckCounter + 1
				rs3.movenext
				loop
			
				rs3.close
				set rs3 = nothing
				response.write "<li><a href='http://172.18.13.31:8081/ShippingTruckFullReport.asp?Job=" & LastJob & "&Floor=" & LastFloor & "' target='_blank'>" & LastJob &  " " & LastFloor & "</a>" 
				if TruckCounter = TotalCounter then
					response.write " " & TruckCounter & " / " & TotalCounter & " (" & Counter & ") "
				else	
					response.write " <font color='red'><b>" & TruckCounter & " / " & TotalCounter & " (" & Counter & ") </b></font>"
				end if
				response.write " - All trucks containing "& LastJob & " " & LastFloor & ": "& TruckNumbers & "</li>"
			end if	
	
rs.movefirst
response.write "<li> Total Windows: " & AllCounter & "</li>"
%>
<li><table border='1'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Description</th><th>Truck ID</th><th>Truck Name</th></tr>
<%
do while not rs.eof
%>
	<tr>
	<td> <%response.write RS("Barcode") %></td>
	<td> <%response.write RS("Job") %></td>
	<td> <%response.write RS("Floor") %></td>
	<td> <%response.write RS("Tag") %></td>
	<td> <%response.write RS("Description") %></td>
	<td> <%response.write RS("Truck") %></td>
	<td> <%response.write TruckName %></td>
	</tr>
<%
	rs.movenext
loop
%>
</table></li>
<%

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
</div>
</body>
</html>
