<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search Production Stock by PO Search page -->
<!--Created May 1st, by Michael Bernholtz at Request of Ruslan Bedoev -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Adding Floor</title>
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
floor=request.querystring("floor")
ParentFloor=request.querystring("Parentfloor")
Jobstatus=request.querystring("JobStatus")
Parent=request.querystring("Parent")
	if Parent= "Parent" then
		IsParent = TRUE
		ParentFloor = Floor
	else
		IsParent = FALSE
		ParentFloor = ParentFloor
	end if
	
		
	
	
AlreadyEntered = False	
	
Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "Select * FROM Z_Floors Where JOB = '" & JOB & "'"
		rs.Cursortype = GetDBCursorTypeInsert
		rs.Locktype = GetDBLockTypeInsert
		rs.Open strSQL, DBConnection
		Response.write "1"
			if JOB <> "" and FLoor <> "" then
				rs.filter = "JOB = '" & JOB & "' and Floor = '" & Floor & "'"
				
				if rs.eof then 
					rs.AddNew
					rs.Fields("Job") = JOB
					rs.Fields("Floor") = Floor
					rs.Fields("ParentFloor") = ParentFloor
					rs.Fields("JobStatus") = JobStatus
					Rs.update

				else 
					AlreadyEntered = TRUE
					Response.write "3"
				end if
			end if

	RS.close
	Set RS = Nothing

DBConnection.close
Set DBCOnnection = nothing

Response.write JOB & " - "
Response.write Floor & " - "
Response.write ParentFloor & " - "
Response.write JObStatus & " - "
%>

<script>
    setTimeout(function(){location.href="AddFloor2.asp?job=<%response.write job%>"} , 100);
</script>

