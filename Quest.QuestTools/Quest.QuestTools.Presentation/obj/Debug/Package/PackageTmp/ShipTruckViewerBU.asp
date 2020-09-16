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
	    <script>

			function check1(TableNum){
			var columnID = "truck" + TableNum;
			var Checker = "trucktab" + TableNum;
			
				if (document.getElementById(Checker).checked == true) {
					document.getElementById(columnID).style.display = "block";
				} else {
					document.getElementById(columnID).style.display = "none";
				};
			}

			function check2(TableNum){
			var columnID = "othertruck" + TableNum;
			var Checker = "othertrucktab" + TableNum;
			
				if (document.getElementById(Checker).checked == true) {
					document.getElementById(columnID).style.display = "block";
				} else {
					document.getElementById(columnID).style.display = "none";
				};
			}

			function check3(TableNum){
			var columnID = "Backorder" + TableNum;
			var Checker = "backordertab" + TableNum;
			
				if (document.getElementById(Checker).checked == true) {
					document.getElementById(columnID).style.display = "block";
				} else {
					document.getElementById(columnID).style.display = "none";
				};
			}

		</script>


    <%
	Truck = Request.QueryString("truck")
	Ticket = Request.QueryString("Ticket")
	Counter = 0
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIP WHERE [Truck] = " & truck & " ORDER BY TAG ASC"
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

	'for i=0 to Ubound(JobsList)
	'	response.write "<td>" & iJob(i) & "</td>"
	'	response.write "<td>" & iFloor(i) & "</td>"
	'next

'rs4 All Shipping Item with Job/Floor in Jobs


	for i=0 to JobLimit
		if WindowWhere = "" then
			WindowWhere = "WHERE ([JOB] = '" & iJob(i) & "' and [Floor] = '" & iFloor(i) & "')"
			HBarWhere = "SELECT JOB,FLOOR,OPENING,HBARQTY FROM Y_Entry  WHERE [JOB] = '" & JOB & "' AND ([Floor] = '" & Floor & "' or [Floor] = '" & Floor2 & "')"
		else
			WindowWhere = WindowWhere & " OR ([JOB] = '" & iJob(i) & "' and [Floor] = '" & iFloor(i) & "')"
			HBArWhere = HBarWhere & " UNION ALL SELECT JOB,FLOOR,OPENING,HBARQTY FROM Y_Entry WHERE [JOB] = '" & JOB & "' AND ([Floor] = '" & Floor & "' or [Floor] = '" & Floor2 & "')"
		end if
	next

strSQL4 = "SELECT * FROM X_SHIP " & WindowWhere
Set rs4 = Server.CreateObject("adodb.recordset")
rs4.Cursortype = 1
rs4.Locktype = 3
if CountryLocation = "USA" then
	rs4.Open strSQL4, DBConnection_Texas
else	
	rs4.Open strSQL4, DBConnection
end if

	for i=0 to Ubound(JobsList)
		if AllWindow = "" then
			AllWindow = "SELECT JOB,FLOOR,TAG FROM " & iJob(i) & "  WHERE  ([Floor] = '" & iFloor(i) & "')"
			AllHBar = "SELECT JOB,FLOOR,OPENING,HBARQTY FROM Y_Entry  WHERE [JOB] = '" & iJob(i) & "' AND ([Floor] = '" & iFloor(i) & "') order by TAG ASC"
		else
			AllWindow = AllWindow & " UNION ALL SELECT JOB,FLOOR,TAG FROM " & iJob(i) & "  WHERE  ( [Floor] = '" & iFloor(i) & "')"
			AllHBar = AllHBar & " UNION ALL SELECT JOB,FLOOR,OPENING,HBARQTY FROM Y_Entry  WHERE [JOB] = '" & iJob(i) & "' AND ([Floor] = '" & iFloor(i) & "') order by TAG ASC"
		end if
	next

strSQL5 = AllWindow
Set rs5 = Server.CreateObject("adodb.recordset")
rs5.Cursortype = 1
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection





%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
<style>
table, th, td {
  padding: 10px;
}
th {
	font-size: 22px
}

</style>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
		
		<%
		Select CASE Ticket
		Case "Open"
		%>
        <a class="button leftButton" type="cancel" href="ShipReportOpen.asp" target="_self">Open</a>
		<%
		Case "Closed"
		%>
        <a class="button leftButton" type="cancel" href="ShipReportClosed.asp" target="_self">Closed</a>
		<%
		Case Else
			if CountryLocation = "USA" then
		%>
			<a class="button leftButton" type="cancel" href="IndexTexas.html#_Ship" target="_self">Scan USA</a>
		<%
			Else
		%>
			<a class="button leftButton" type="cancel" href="Index.html#_Ship" target="_self">Scan Canada</a>
		<% 
			End if 
		End Select
		%>
        </div>
   
 
        <ul id="Profiles" title="Window Report" selected="true">
		<%
		if CountryLocation = "USA" then
		else
		%>
        <li><a href= "ShippingTruckViewerExcel.asp?Truck=<%response.write Truck %>&ticket=<%response.write Ticket %>" target="_self" >Send to Excel</a></li>
		<%
		end if%>
		<li>Truck Name: <% response.write RS2("TruckName")%></li>
		<li>sList: <% response.write RS2("sList")%></li>
			<% 
			do while not rs.eof
				if rs("Window") = "Window" or rs("Window") = "H-Bar" then 
					Counter = Counter + 1
				end if
			rs.movenext
			loop
			If Counter > 0 Then rs.movefirst
			response.write "<li> Total Items   " & Counter & "</li>"
			%>
			<table border="0">
			<tr><TH>Items on truck</TH><TH>Items on other trucks in sList</TH><TH>Back Order</TH></tr>
			<tr><td valign="Top">
<%
response.write "<table border='1' class='sortable'>"
rs4.filter = "TRUCK = " & truck
Jobcount = 0
ChangeCounter = 0
PreviousJob = "0"
PreviousFloor = "0"
TotalCount = rs4.RecordCount
do while not rs4.eof
	if NOT(PreviousJob = RS4("Job") AND  PreviousFloor = RS4("Floor")) then
	ChangeCounter = ChangeCounter + 1
		if PreviousJob = "0" then 
		else
			response.write "</tbody>"
			response.write "<tbody>"
			response.write "<tr>"
			response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
			response.write "<td><b>" & JobCount & " / " & TotalCount & "</b></td>"
			response.write "</tr>"
			response.write "</tbody>"
		end if
		
		response.write "<tr>"
		response.write "<td><b><big>" & RS4("Job")& RS4("Floor") & "</big></b>"
		response.write "<input type='checkbox' name='truck' id ='trucktab" & ChangeCounter& "' onclick='check1(" & ChangeCounter & ")' />"
		response.write "</td></tr>"
		response.write "<tbody class='hide' id='truck" & ChangeCounter & "' style='display:none' >"
		response.write "<tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th><th>Description</th></tr>"
	end if

	response.write "<tr>"
	response.write "<td>" & RS4("Barcode") & "</td>"
	response.write "<td>" & RS4("Job") & "</td>"
	response.write "<td>" & RS4("Floor") & "</td>"
	response.write "<td>" & RS4("TAG") & "</td>"
	response.write "<td>" & RS4("SHIPDATE") & "</td>"
	response.write "<td>" & RS4("Description") &  "</td>"	
	response.write "</tr>"
		Jobcount = Jobcount + 1
		PreviousJob = RS4("Job")
		PreviousFloor = RS4("Floor")

	rs4.movenext
loop

	response.write "</tbody>"
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
	response.write "<td><b>" & JobCount & "/" & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"
	
	response.write "</tbody>"
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"
response.write "</table>"
	



%>
</td><td valign="Top">
<%
response.write "<table border='1' class='sortable'>"
rs4.filter = "TRUCK <> " & truck
JobCount = 0
ChangeCounter = 0
PreviousJob = "0"
PreviousFloor = "0"
TotalCount = rs4.RecordCount
do while not rs4.eof


	if NOT(PreviousJob = RS4("Job") AND  PreviousFloor = RS4("Floor")) then
	ChangeCounter = ChangeCounter + 1
		if PreviousJob = "0" then 
		else
			response.write "</tbody>"
			response.write "<tbody>"
			response.write "<tr>"
			response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
			response.write "<td><b>" & JobCount & " / " & TotalCount & "</b></td>"
			response.write "</tr>"
			response.write "</tbody>"
		end if
		
		response.write "<tr>"
		response.write "<td><b><big>" & RS4("Job")& RS4("Floor") & "</big></b>"
		response.write "<input type='checkbox' name='othertruck' id ='othertrucktab" & ChangeCounter& "' onclick='check2(" & ChangeCounter & ")' />"
		response.write "</td></tr>"
		response.write "<tbody class='hide' id='othertruck" & ChangeCounter & "' style='display:none' >"
		response.write "<tr><th>Truck</th><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Scan Date</th><th>Description</th></tr>"
	end if

	response.write "<tr>"
	response.write "<td><b>" & RS4("Truck") & "</b></td>"
	response.write "<td>" & RS4("Barcode") & "</td>"
	response.write "<td>" & RS4("Job") & "</td>"
	response.write "<td>" & RS4("Floor") & "</td>"
	response.write "<td>" & RS4("TAG") & "</td>"
	response.write "<td>" & RS4("SHIPDATE") & "</td>"
	response.write "<td>" & RS4("Description") & "</td>"	
	response.write "</tr>"
		Jobcount = Jobcount + 1
		PreviousJob = RS4("Job")
		PreviousFloor = RS4("Floor")

	rs4.movenext
loop

	response.write "</tbody>"
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
	response.write "<td><b>" & JobCount & "/" & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"
	
	response.write "</tbody>"
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"


response.write "</table>"
%>

</td><td valign="Top">
<%
response.write "<table border='1' class='sortable'>"
rs4.filter =""
JobCount = 0
ListCount = 0
ChangeCounter = 0
PreviousJob = "0"
PreviousFloor = "0"
TotalCount = rs5.RecordCount
do while not rs5.eof


	BarcodeTest = RS5("Job") & RS5("Floor") & RS5("TAG")
	rs4.filter = " Barcode = '" & BarcodeTest & "'"
	if rs4.eof then
		ListCount = ListCount + 1
		
		if NOT(PreviousJob = RS5("Job") AND  PreviousFloor = RS5("Floor")) then
		ChangeCounter = ChangeCounter + 1
			if PreviousJob = "0" then 
			else
				response.write "</tbody>"
				response.write "<tbody>"
				response.write "<tr>"
				response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
				response.write "<td><b>" & JobCount & " / " & TotalCount & "</b></td>"
				response.write "</tr>"
				response.write "</tbody>"
			end if
			
			response.write "<tr>"
			response.write "<td><b><big>" & RS5("Job")& RS5("Floor") & "</big></b>"
			response.write "<input type='checkbox' name='Backorder' id ='backordertab" & ChangeCounter& "' onclick='check3(" & ChangeCounter & ")' />"
			response.write "</td></tr>"
			response.write "<tbody class='hide' id='Backorder" & ChangeCounter & "' style='display:none' >"
			response.write "<tr><th>Job</th><th>Floor</th><th>Tag</th></tr>"
			JobCount = 0
		end if
			
			response.write "<tr>"
			response.write "<td>" & RS5("Job") & "</td>"
			response.write "<td>" & RS5("Floor") & "</td>"
			response.write "<td>" & RS5("TAG") & "</td>"	
			response.write "</tr>"
			JobCount = JobCount + 1
		PreviousJob = RS5("Job")
		PreviousFloor = RS5("Floor")
	
	end if
	
	rs5.movenext
loop
	response.write "</tbody>"
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>" & PreviousJob & PreviousFloor & " Total:</b></td>"
	response.write "<td><b>" & JobCount & "/" & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"
	
	response.write "<tbody>"
	response.write "<tr>"
	response.write "<td><b>Total</b></td>"
	response.write "<td><b>" & ListCount & " / " & TotalCount & "</b></td>"
	response.write "</tr>"
	response.write "</tbody>"
	
	
response.write "</table>"
%>
</td><tr></table></li>


<%
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing
rs4.close
set rs4 = nothing
rs5.close
set rs5 = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
               
    </ul>      
  
</body>
</html>
