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
<!-- July 2019 New Format for sLIst Trucks-->
<!-- Date: October 9, 2019
	Modified By: Michelle Dungo
	Changes: Modified to remove limit for searching from top 50 to include all
-->

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
	
<!-- DataTables CSS -->
	<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
	<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
	<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

	<script type="text/javascript">
		$(document).ready( function () {
		$('#Job').DataTable({
			"iDisplayLength": 25
		});
	});
  
	</script>
	    


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
 <div class="toolbar">
        <h1 id="pageTitle">Closed Trucks</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "index.html#_Ship"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Ship"
				HomeSiteSuffix = ""
			end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
   
 
        <ul id="Profiles" title="Closed Trucks" selected="true">
		<%
		if CountryLocation = "USA" then
		else
		%>
        <li><a href= "ShipReportClosedExcel.asp" target="_self" >Send Closed List to Excel</a></li>
		<%
		end if%>
        <li>Shipping Closed Trucks (50 Most Recent)</li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
response.write "<li><table border='1' id='Job'><THead><tr><th>Truck Name</th><th>System Number</th><th>Jobs/Floors</th><th>Open Date</th><th>Closed Date</th><th>Item count</th><th>Backorder count</th><th>View</th></tr></THEAD><TBODY>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("truckName") & "</td>"
	response.write "<td>" & RS("ID") & "</td>"
	response.write "<td style='word-break:break-all;'>" & RS("sList") & "</td>"
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
	
	response.write "<td><a class='greenButton' href='ShipTruckViewer.asp?truck=" & RS("ID") & "&Ticket=Closed' target='_self' >View All Items </a></td>"
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
