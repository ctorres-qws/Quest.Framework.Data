<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
	<!-- Add Floor Function adds Floor information to the Z_Floors Table -->
	<!-- First Checks Job Status of Guaranteed / Measured / Mixed -- if Mixed, forces choice by floor -->
	<!-- Continues to allow new floors to be entered as Parent or Child -->
	<!-- Michael Bernholtz, October 2018 - for Ariel Aziza, David Ofir and Jody Cash -->
				
<!-- Add Floor Page 3 of 3 -->
		<!-- Select Job Status for All FLoors on the Job that do not have a Status-->	
		<!-- Add Floor Status information - only for Jobs  -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Adding Floor Status</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  
  
</head>
<body>
<font face="Arial" size="2">
Floor Entry Submitting, Please wait:
<br>
</font>
<br>
<%
job=request.querystring("job")	
FloorStatus = ""
	
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_Floors Where JOB = '" & JOB & "' ORDER BY Floor ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection  

	Do while Not rs.eof
	FloorID = rs("ID") & ""
	FloorStatus = request.Querystring(FloorID)
	response.write FloorStatus
			if FloorStatus = "Guaranteed" or FloorStatus = "Measured" then								
				
				rs.Fields("JobStatus") = FloorStatus
				Rs.update
			end if
	
	rs.movenext
	loop
	
	
	RS.close
	Set RS = Nothing

DBConnection.close
Set DBCOnnection = nothing


%>

<script>
    setTimeout(function(){location.href="AddFloor3.asp?job=<%response.write job%>"} , 100);
</script>

