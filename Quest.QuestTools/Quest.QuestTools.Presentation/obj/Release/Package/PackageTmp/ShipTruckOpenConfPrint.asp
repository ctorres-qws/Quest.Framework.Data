<html>
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<head>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<title>Window Label</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">

body,td,th {
	font-family: Verdana, Geneva, sans-serif;
	font-weight: bold;
	font-size: 200px;
}

@media print
{
table {page-break-after:always}
}

</style>

</head>
<body  align= "center" valign ="middle" link="#000000" vlink="#C0C0C0" alink="#F6F000">
<%
Truck = Request.Querystring("Truck")
Set rsTruck = Server.CreateObject("adodb.recordset")
strSQL = "SELECT sList from X_SHIP_TRUCK where ID = " & TRUCK 
rsTruck.Cursortype = 1
rsTruck.Locktype = 3
if CountryLocation = "USA" then
	rsTruck.Open strSQL, DBConnection_Texas
else	
	rsTruck.Open strSQL, DBConnection
end if

FloorList = rsTruck("sList")

rsTruck.close
set rsTruck = Nothing

JobsList = Split(FloorList, ",")
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
	if FloorList ="" then 
		JobLimit = -1
	else
		JobLimit = 0
		SplitItem = FloorList
		iJob(0) = Left(SplitItem,3)
		iFloor(0) = Right(SplitItem,(Len(SplitItem)-3))
	end if 
end if

%>
<%

for i=0 to JobLimit

Set rsJob = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From [" & iJob(i) & "] where Floor = '" & iFloor(i) & "'"
rsJob.Cursortype = 1
rsJob.Locktype = 3
rsJob.Open strSQL, DBConnection
WindowCount = rsJob.RecordCount
rsJob.close
set rsJob = Nothing
%>



<div>
<table align= "center" valign ="middle" frame="box"  cellspacing="1" cellpadding="1">

	<tr>
		<TD align = 'center' valign = 'center'> <%Response.write iJob(i)%></TD>
	</tr>
	<tr>
		<TD align = 'center' valign = 'center'><%Response.write iFloor(i)%> </TD>
	</tr>
	<tr>
		<TD align = 'center' valign = 'center'><%Response.write WindowCount%>W </TD>
	</tr>

</table>
</Div>


<%
next
%>
</body>

<%

DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing
%>

</HTML>


