<!-- https://www.w3schools.com/w3css/w3css_progressbar.asp -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<!--Shipping Monitor Report - November 5,2019, Michael Bernholtz based on design by Pranav Gulavene and Arjon -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping View</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<meta http-equiv="refresh" content="100" >
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    


    <%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT Top 10 * FROM X_SHIP_TRUCK WHERE [Active] = True ORDER BY DOCKNUM ASC")
rs.Cursortype = 2
rs.Locktype = 3

if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if
%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle">Ship Monitor</h1>
             
    </div>
   
 
        <ul id="Profiles" title="Open Trucks" selected="true">

        <li>Open Trucks: <%Response.write Now %></li>

	
<% 
response.write "<li><table border='1' id='Job' width ='100%'><THead><tr><th>Dock #</th><th>Truck Name</th><th>Jobs/Floors</th><th>Open Date</th><th>Scanned</th><th> Missing</th><TH>STATUS BAR (Percent Loaded)</TH></tr></THEAD><TBODY>"
do while not rs.eof
	response.write "<tr>"
	response.write "<TD style='font-size: 165%;'><B>" & RS("DockNum") & "</B></td>"
	response.write "<TD><B>" & RS("truckName") & "</B></td>"
	response.write "<td style='word-break:break-all; font-size: 165%;'><B>" & RS("sList") & "</B></td>"
	response.write "<TD style='font-size: 115%;'><B>" & RS("CreateDate") & "</B></td>"
	
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
	
	response.write "<TD style='font-size: 300%;'>" & Counter & "</td>"
	
	
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
	response.write "<TD style='font-size: 300%;'>" & BackOrderCounter & "</td>"
	
	ProgressIn = Counter
	ProgressOut = Counter + BackOrderCounter
	ProgressValue = Round((ProgressIn/ProgressOut)*100,0)
	
	response.write "<td>"
	If (ProgressIn/ProgressOut = 1) Then
		%>
		<div class="w3-light-grey w3-round">
		<div class="w3-container w3-round w3-green w3-center" style="width:<%Response.write ProgressValue%>%"><%Response.write ProgressValue%>%</div>
		</div>
		<%
	ElseIf (ProgressIn/ProgressOut > 0.75) Then
			%>
		<div class="w3-light-grey w3-round">
		<div class="w3-container w3-round w3-blue w3-center" style="width:<%Response.write ProgressValue%>%"><%Response.write ProgressValue%>%</div>
		</div>
		<%
	ElseIf (ProgressIn/ProgressOut > 0.40) Then
			%>
		<div class="w3-light-grey w3-round">
		<div class="w3-container w3-round w3-orange w3-center" style="width:<%Response.write ProgressValue%>%"><%Response.write ProgressValue%>%</div>
		</div>
		<%
	Else
		%>
		<div class="w3-light-grey w3-round">
		<div class="w3-container w3-round w3-red w3-center" style="width:<%Response.write ProgressValue%>%"><%Response.write ProgressValue%>%</div>
		</div>
		<%
	End if
	
	response.write "</td>"
	response.write "</tr>"
	rs.movenext
loop


rs.close
set rs = nothing


%>
</TBody></Table>
</LI>
<li>Last 5 Scans</li>
<LI>
<TABLE border = '1' width='100%'>
<THead><TR><TH>Barcode</TH><TH>Scan Time</TH><TH>DOCK</TH></TR></THead>
<TBody>
<%
Set rs3 = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT Top 100 * FROM X_SHIP_TRUCK ORDER BY ID DESC")
rs3.Cursortype = 2
rs3.Locktype = 3

if CountryLocation = "USA" then
	rs3.Open strSQL, DBConnection_Texas
else	
	rs3.Open strSQL, DBConnection
end if

Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT TOP 5 * FROM X_SHIP ORDER BY ID DESC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	if CountryLocation = "USA" then
		rs2.Open strSQL2, DBConnection_Texas
	else	
		rs2.Open strSQL2, DBConnection
	end if
	do while not rs2.eof
		Response.write "<TR>"
		Response.write "<TD style='font-size: 200%;'>" & rs2("JOB") & rs2("FLOOR") & "-" & rs2("TAG") & "</TD>"
		Response.write "<TD style='font-size: 200%;'>" & rs2("ShipDate") & " - " & rs2("ShipTime") & "</TD>"
		
		rs3.filter =""
		rs3.Filter = "ID = " & rs2("TRUCK")
		Response.write "<TD style='font-size: 200%;'>" & rs3("DockNum") & "</TD>"
		Response.write "</TR>"
	rs2.movenext
	loop
%>
</TBody>




<%
rs3.close
set rs3 = nothing

	rs2.close
	set rs2 = nothing
	
	DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing
%>
</TABLE>
</LI>               
    </ul>      
  
</body>
</html>
