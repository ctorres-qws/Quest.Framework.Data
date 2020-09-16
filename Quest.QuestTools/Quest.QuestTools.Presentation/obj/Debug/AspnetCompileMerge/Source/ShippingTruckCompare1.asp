<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Shipping Table Reporting - December 2015, Michael Bernholtz -->
<!-- At the request of Jody Cash and Alex Sofienko, this tool allows viewing items on Trucks, Active, Closed, and comparing to Job Expectation -->
<!-- At the request of Sokol monerolli and Jody Cash, this tool now checks for Parent Job and shows ALL missing per Total job. March 2018 -->



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
	Ticket = Request.QueryString("Ticket")
%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
		
		<%
		Select CASE Ticket
		Case "Open"
		%>
        <a class="button leftButton" type="cancel" href="ShippingReportOpen.asp" target="_self">Open</a>
		<%
		Case "Closed"
		%>
        <a class="button leftButton" type="cancel" href="ShippingReportClosed.asp" target="_self">Closed</a>
		<%
		Case Else
		%>
        <a class="button leftButton" type="cancel" href="Index.html#_ShipReport" target="_self">Ship Home</a>
		<%
		End Select
		%>
        </div>
        <ul id="Profiles" title="Window Report" selected="true">
        <li><a href= "ShippingTruckViewerExcel.asp?Truck=<%response.write Truck %>&ticket=<%response.write Ticket %>" target="_self" >Send to Excel</a></li>
		<li> Click on the Headers of each column to sort Ascending/Descending</li>
	
<% 
	Job = Request.QueryString("Job")
	Floor = Request.QueryString("Floor")
	Counter = 0
	MissingCounter = 0
	
	' Addition for Parent Job 
	'Step 1 connect to Z_JObs and check the Parent Job.
	'Step Two display for JOB of all the tables not just 
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT JOB,PARENT FROM Z_JOBS WHERE [JOB] = '" & JOB & "'"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL3, DBConnection
	
	ParentJob = RS3("PARENT")
	rs3.close
	set rs3 = nothing
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT JOB FROM Z_JOBS WHERE [PARENT] = '" & ParentJob & "'"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL3, DBConnection
	
	strSQL2 = ""
	JobList = ""
	JobSQL = ""
	
	if rs3.eof then 
		strSQL2 = "SELECT * FROM " & JOB  & " WHERE [Floor] = '" & Floor & "'"
	else
		Do while not rs3.eof
			
			Set rs4 = Server.CreateObject("adodb.recordset")
			strSQL4 = "SELECT TOP 1 * FROM [" & RS3("JOB") & "]"
			rs4.Cursortype = 2   
			rs4.Locktype = 3
			On Error Resume Next
			rs4.Open strSQL4, DBConnection
			
			
		
			if Err.Number <> 0 then
					Errlist = Errlist & " " & rs3("JOB")
					Err.clear
			else
					JobList = JobList & " " & rs3("JOB")
					if JOBSQL = "" then
					JOBSQL = "[JOB] = '" & rs3("Job") & "'" 
					else
					JOBSQL = JOBSQL & " or [JOB] = '" & rs3("Job") & "'" 
					End if
					Err.clear
					
					if strSQL2 = "" then
						strSQL2 = "SELECT JOB,FLOOR,TAG FROM " & RS3("JOB")  & " WHERE [Floor] = '" & Floor & "'"
						
					else
						strSQL2 = strSQL2 & " UNION ALL SELECT JOB,FLOOR,TAG FROM " & RS3("JOB")  & " WHERE [Floor] = '" & Floor & "'"
					end if
			end if

		On Error GOTO 0
		If Not (rs4 Is Nothing) Then
			Set rs4 = Nothing
		End If
		
		rs3.movenext
		loop
	rs3.close
	set rs3 = nothing
	
	end if
	strSQL2 = strSQL2 & " ORDER BY JOB ASC, TAG ASC"
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_SHIPPING WHERE (" & JOBSQL & ") AND [Floor] = '" & Floor & "' ORDER BY TAG ASC"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection

Set rs = GetDisconnectedRS(strSQL, DBConnection)

Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL2 = "SELECT * FROM " & JOB  & " WHERE [Floor] = '" & Floor & "' ORDER BY TAG ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection


do while not rs.eof
Counter = Counter + 1
rs.movenext
loop
If Counter > 0 Then rs.movefirst
response.write "<li>Included Jobs: " & Joblist & "</li>"
response.write "<li> Missing Windows</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Status</th></tr>"

Do while not RS2.eof
if Left(RS2("Tag"),1) = "-" then
JobTag = Right(RS2("TAg"),Len(RS2("TAg"))-1)
else
JobTag =RS2("Tag")
End if
rs.filter = "TAG = '" & JobTag & "'"
if rs.eof then
TruckTag = "Empty"
else
TruckTag = RS("Tag")
end if
if JobTag <> TruckTag then
MissingCounter = MissingCounter + 1

response.write "<tr>"
	response.write "<td>" & RS2("Job") & RS2("Floor") &  RS2("Tag") & "</td>"
	response.write "<td>" & RS2("Job") & "</td>"
	response.write "<td>" & RS2("Floor") & "</td>"
	response.write "<td>" & RS2("TAG") & "</td>"
	response.write "<td> Missing </td>"
	response.write "</tr>"

end if
rs2.movenext
loop


response.write "</table></li>"
response.write "<li> Total Windows Missing " & MissingCounter & "</li>"

response.write "<li> Total Windows Scanned for the Floor " & Counter & "</li>"


response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Job</th><th>Floor</th><th>Tag</th><th>Truck</th><th>Ship Date</th><th>Description</th></tr>"
rs.filter = ""
rs.movefirst
Do while not rs.eof
response.write "<tr>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("Truck") & "</td>"
	
	Set rsTruck = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT top 1 * FROM X_SHIPPING_TRUCK  WHERE [ID] = " & RS("TRUCK")
	rsTruck.Cursortype = 2
	rsTruck.Locktype = 3
	rsTruck.Open strSQL, DBConnection
	response.write "<td>" & rsTruck("SHipDate") & "</td>"
	rsTruck.Close
	Set rsTruck = nothing
	
	
	response.write "<td>" & RS("Description") & "</td>"
	response.write "</tr>"
rs.movenext
Loop


Response.write "</table></li>"
response.write "<li>Jobs with No windows currently in System: " & Errlist & "</li>"


rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>      
  
</body>
</html>
