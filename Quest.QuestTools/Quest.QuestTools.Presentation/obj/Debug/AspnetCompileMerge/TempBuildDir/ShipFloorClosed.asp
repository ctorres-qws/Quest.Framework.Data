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

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping Floor View</title>
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
   
 
        <ul id="Profiles" title="Closed Floors" selected="true">
        <li>Floors Available on Closed Trucks</li>

	
<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIP_TRUCK WHERE [Active] = FALSE ORDER BY SHIPDATE DESC")
rs.Cursortype = 2
rs.Locktype = 3
if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if

ActiveFloorList = ""
Do While Not rs.eof

	sList = rs("sList")
	JobsList = Split(sList, ",")
	JobLimit = Ubound(JobsList)
	if sList ="" then 
		JobLimit = -1
	end if
	i = 0
	if (JobLimit > 0) Then
		Do Until i > Joblimit
			JF = Trim(Jobslist(i))
			if instr(ActiveFloorList,JF) > 0 Then
			else 
				if Len(ActiveFloorList) = 0  then
					ActiveFloorList = ActiveFloorList & JF
				else
					ActiveFloorList = ActiveFloorList & "," & JF
				end if
			end if
			i = i+1
		loop
	else
		
		if JobLimit = 0 then
			JF = sList
			if instr(ActiveFloorList,JF) > 0 Then
			else 
				if Len(ActiveFloorList) = 0  then
					ActiveFloorList = ActiveFloorList & JF
				else
					ActiveFloorList = ActiveFloorList & "," & JF
				end if
			end if
		end if 
	end if
rs.movenext
loop




response.write "<li><table border='1' id='Job'><THead><tr><th>Job</th><th>Floor</th><th>Scanned</th><th>Total Windows</th><th>View</th></tr></THead><TBody>"
	
	ActiveFloorDisplay = Split(ActiveFloorList, ",")
	JobLimit = Ubound(ActiveFloorDisplay)
	

For i=0 to JobLimit  
   For j=i to JobLimit  
      if ActiveFloorDisplay(i)>ActiveFloorDisplay(j) then 
          TemporalVariable=ActiveFloorDisplay(i) 
          ActiveFloorDisplay(i)=ActiveFloorDisplay(j) 
          ActiveFloorDisplay(j)=TemporalVariable 
     end if 
   next  
next 	
	

i= 0
do Until i> Joblimit

	JF = Trim(ActiveFloorDisplay(i))
	JOB = Left(JF,3)
	FLOOR = Right(JF, Len(JF)-3)


	response.write "<tr>"
	response.write "<td>" & JOB & "</td>"
	response.write "<td>" & FLOOR & "</td>"

	Counter = 0
	set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM X_SHIP WHERE [DELETED] = FALSE AND [JOB] = '" & JOB & "' AND [FLOOR] = '" & FLOOR & "' ORDER BY TAG ASC"

	rs2.Cursortype = 1
	rs2.Locktype = 3
	
	if CountryLocation = "USA" then
		rs2.Open strSQL2, DBConnection_Texas
	else	
		rs2.Open strSQL2, DBConnection
	end if
	counter = rs2.RecordCount
	rs2.close
	set rs2 = nothing
	
	response.write "<td>" & Counter & "</td>"
	
	AllCounter = 0
	set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT * FROM [" & JOB & "] WHERE [FLOOR] = '" & FLOOR & "' ORDER BY TAG ASC"

	rs3.Cursortype = 1
	rs3.Locktype = 3	
	rs3.Open strSQL3, DBConnection
	Allcounter = rs3.RecordCount
	rs3.close
	set rs3 = nothing
	
	response.write "<td>" & AllCounter & "</td>"
	
	
	
	
	response.write "<td><a class='greenButton' href='ShipFloorViewer.asp?Job=" & JOB & "&Floor=" & FLOOR & "&ticket=close' target='_self' >View All Items </a></td>"
	response.write "</tr>"

	
i = i+1 
loop


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
	</TBody></Table></li>
    </ul>      
  
</body>
</html>
