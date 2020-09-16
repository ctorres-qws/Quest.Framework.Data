<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!--Excel version of Report is a copy of ShipReportClosed requested by Pranav Gulavane October 2019  -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping View</title>
  

    <%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT Top 50 * FROM X_SHIP_TRUCK WHERE [Active] = False ORDER BY ID DESC")
rs.Cursortype = 2
rs.Locktype = 3

if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if
%>
    </head>
<body>
 <%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=ClosedTruckData.xls"
%>
 
        <ul id="Profiles" title="Closed Trucks" selected="true">
        <li>Shipping Closed Trucks (50 Most Recent) <%response.write Date%></li>
	
<% 
response.write "<li><table border='1' id='Job'><THead><tr><th>Truck Name</th><th>System Number</th><th>Jobs/Floors</th><th>Open Date</th><th>Closed Date</th><th>Item count</th><th>Backorder count</th></tr></THEAD><TBODY>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("truckName") & "</td>"
	response.write "<td>" & RS("ID") & "</td>"
	response.write "<td>" & RS("sList") & "</td>"
	response.write "<td>" & RS("CreateDate") & "</td>"
	response.write "<td>" & RS("ShipDate") & "</td>"
	
	Counter = 0
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND [Truck] = " & RS("ID") & " ORDER BY TAG ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	if CountryLocation = "USA" then
		rs2.Open strSQL2, DBConnection_Texas
	else	
		rs2.Open strSQL2, DBConnection
	end if
	do while not rs2.eof
			Counter = Counter + 1
	rs2.movenext
	loop
	rs("Itemcount") = Counter
	rs.update
	rs2.close
	set rs2 = nothing
	
	response.write "<td>" & Counter & "</td>"
	
	
	BackOrderCounter = 0
	
sList = rs("sList")
JobsList = Split(sList, ",")
Dim iJob(25)
Dim iFloor(25)
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
	
		do while not rs5.eof
			BarcodeTest = RS5("Job") & RS5("Floor") & RS5("TAG")
			rs4.filter = " Barcode = '" & BarcodeTest & "'"
			if rs4.eof then
				BackOrderCounter = BackOrderCounter + 1
			end if
	
		rs5.movenext
		loop
		
		rs4.close
		rs5.close
		set rs4 = nothing
		set rs5 = nothing 
	next 
	response.write "<td>" & BackOrderCounter & "</td>"
	response.write "</tr>"
	rs.movenext
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
</TBody></Table>
</LI>
               
    </ul>      
  
</body>
</html>
