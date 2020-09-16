<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->
<!-- May 2019 - Updated to include Texas Database-->
<!-- July 2019 - Adding Collapsable tables And Close Button-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style type="text/css">
input.largerCheckbox
{
	width: 25px;
	height: 25px;
}
</style>



    <%
	Truck = Request.QueryString("truck")
	Ticket = Request.QueryString("Ticket")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND [Truck] = " & truck & " ORDER BY TAG ASC"
rs.Cursortype = 2
rs.Locktype = 3
if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT top 1 * FROM X_SHIP_TRUCK WHERE [ID] = " & Truck 
rs2.Cursortype = 2
rs2.Locktype = 3
if CountryLocation = "USA" then
	rs2.Open strSQL2, DBConnection_Texas
else	
	rs2.Open strSQL2, DBConnection
end if

sList = rs2("sList")
JobsList = Split(sList, ",")
Dim iJob(50)
Dim iFloor(50)
JobLimit = Ubound(JobsList)

if (JobLimit => 1) Then 
    for i=0 to Ubound(JobsList)
		SplitItem = Trim(Jobslist(i))
		iJob(i) = Left(SplitItem,3)
		iFloor(i) = Right(SplitItem,(Len(SplitItem)-3))
		test = "7"
    next
else
	if sList ="" then 
		JobLimit = -1
	else
		JobLimit = 0
		SplitItem = sList
		iJob(0) = Left(SplitItem,3)
		iFloor(0) = Right(SplitItem,(Len(SplitItem)-3))
	end if 
end if


%>
    </head>
<body>
<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=ScannedTruckData.xls"
%>
 
        <ul id="Profiles" title="Window Report" selected="true">
		<li>Truck Name: <% response.write RS2("TruckName")%></li>
		<li>Floors on Truck: <% response.write RS2("sList")%></li>
			<% 
			do while not rs.eof
				if rs("Window") = "Window" or rs("Window") = "H-Bar" then 
					Counter = Counter + 1
				end if
			rs.movenext
			loop
			If Counter > 0 Then rs.movefirst
			response.write "<li> Total Items on this Truck   " & Counter & "</li>"

			if ticket = "Closed" then
			%>
			 <li> Truck Closed on <% Response.write rs2("Shipdate") %></li>
			<%
			end if 
			%>
			<table border="0">
			<tr><TH align='left'>Items on this truck</TH><TH align='left'>Items on other truck(s)</TH><TH align='left'>Back Order</TH></tr>

<%
	ChangeCounter= 0
	for i=0 to JobLimit
		strSQL4 = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND ([JOB] = '" & iJob(i) & "' and [Floor] = '" & iFloor(i) & "')"
		strSQL5 = "SELECT JOB,FLOOR,TAG FROM " & iJob(i) & "  WHERE  ([Floor] = '" & iFloor(i) & "') ORDER BY TAG ASC"
	
Set rs4 = Server.CreateObject("adodb.recordset")
rs4.Cursortype = 1
rs4.Locktype = 3
if CountryLocation = "USA" then
	rs4.Open strSQL4, DBConnection_Texas
else	
	rs4.Open strSQL4, DBConnection
end if

Set rs5 = Server.CreateObject("adodb.recordset")
rs5.Cursortype = 1
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection			
			
			
%>		
			<tr><td valign="Top">
<%			
			
response.write "<table border='1' class='sortable'>"
rs4.filter = "TRUCK = " & truck
TotalCount = rs4.RecordCount
OECount = rs5.RecordCount
	response.write "<tr>"
	response.write "<td><big><b>" & iJob(i)& iFloor(i) & "</b> (" & TotalCount & " / " & OECount & ") </big>"
	response.write "</td></tr>"
	response.write "<tbody class='hide' id='truck" & ChangeCounter & "' style='display:none' >"
	response.write "<tr><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th></tr>"

	do while not rs4.eof		
	
		response.write "<tr>"
		response.write "<td>" & RS4("Job") & "</td>"
		response.write "<td>" & RS4("Floor") & "</td>"
		response.write "<td>" & RS4("TAG") & "</td>"
		response.write "<td>" & RS4("SHIPDATE") & " " & RS4("SHIPTIME") & "</td>"	
		response.write "</tr>"

	rs4.movenext
	loop

	response.write "</tbody>"
response.write "</table>"
	



%>
</td><td valign="Top">
<%
response.write "<table border='1' class='sortable'>"
rs4.filter = "TRUCK <> " & truck
JobCount = 0
TotalCount = rs4.RecordCount
OECount = rs5.RecordCount

	response.write "<tr>"
	response.write "<td><big><b>" & iJob(i)& iFloor(i) & "</b> (" & TotalCount & " / " & OECount & ") </big>"
	response.write "</td></tr>"
	response.write "<tbody class='hide' id='othertruck" & ChangeCounter & "' style='display:none' >"
	response.write "<tr><th>Truck</th><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th></tr>"

	do while not rs4.eof
		response.write "<tr>"
		response.write "<td><b>" & RS4("Truck") & "</b></td>"
		response.write "<td>" & RS4("Job") & "</td>"
		response.write "<td>" & RS4("Floor") & "</td>"
		response.write "<td>" & RS4("TAG") & "</td>"
		response.write "<td>" & RS4("SHIPDATE") & " " & RS4("SHIPTIME") & "</td>"
		response.write "</tr>"
		Jobcount = Jobcount + 1
		rs4.movenext
	loop

	response.write "</tbody>"

response.write "</table>"
%>

</td><td valign="Top">
<%
response.write "<table border='1' class='sortable'>"
rs4.filter =""
JobCount = 0
TotalCount = rs5.RecordCount

	if not rs5.eof then
		rs5.movefirst	
		do while not rs5.eof
			BarcodeTest = RS5("Job") & RS5("Floor") & RS5("TAG")
			rs4.filter = " Barcode = '" & BarcodeTest & "'"
			if rs4.eof then
				JobCount = JobCount + 1
			end if
		rs5.movenext
		loop
		rs5.Movefirst
	end if


	response.write "<tr>"
	response.write "<td><big><b>" & iJob(i)& iFloor(i) & "</b> (" & JobCount & " / " & OECount & ") </big>"
	response.write "</td></tr>"
	response.write "<tbody class='hide' id='Backorder" & ChangeCounter & "' style='display:none' >"
	response.write "<tr><th>Job</th><th>Floor</th><th width = '100px'>Tag</th></tr>"
	
	do while not rs5.eof
	BarcodeTest = RS5("Job") & RS5("Floor") & RS5("TAG")
	rs4.filter = " Barcode = '" & BarcodeTest & "'"
	if rs4.eof then
			response.write "<tr>"
			response.write "<td>" & RS5("Job") & "</td>"
			response.write "<td>" & RS5("Floor") & "</td>"
			response.write "<td>" & Replace(RS5("TAG"),"-","") & "</td>"	
			response.write "</tr>"
	end if
	
	rs5.movenext
loop
	response.write "</tbody>"
	
response.write "</table>"

rs4.close
set rs4 = nothing
rs5.close
set rs5 = nothing
ChangeCounter = ChangeCounter + 1

response.write "<tr><td colspan ='3'>"
response.write "<HR WIDTH='60%' ALIGN='LEFT'><HR WIDTH='60%' ALIGN='RIGHT'><HR WIDTH='60%' ALIGN='LEFT'><HR WIDTH='60%' ALIGN='RIGHT'>"
response.write "</td></tr>"
next

%>
</td></tr></table></li>

<%
if ticket = "Open" then
%>
 <li><a class='redButton' href="ShipTruckCloseConf.asp?Truck=<%response.write Truck%>" target="_self" >Close This Truck!</a></li>
<%
end if 
%>






<%
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing

DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
               
    </ul>      
  
</body>
</html>
