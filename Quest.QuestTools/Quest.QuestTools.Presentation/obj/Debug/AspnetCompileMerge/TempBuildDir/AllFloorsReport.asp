<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	
<!-- October 2018 - Michael Bernholtz - Floor Entry Page -->
<!-- Created for Jody Cash, David Ofir, Ariel Aziza initially for Jamb Receptor Entry -->
<!-- Future use will include Statement of Value and Area/SQFT calculations -->

<!-- Page writes to Z_Floors -->

		 
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Floor Entry</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
    </head>
<body>

<%
AlreadyEntered = False
JobStatusRequired = False
JOB = REQUEST.QueryString("Job")
FLOOR = REQUEST.QueryString("Floor")
PARENTFLOOR = REQUEST.QueryString("ParentFloor")
JOBSTATUS = REQUEST.QueryString("JobStatus")
if JOBSTATUS = "-" then
	JobStatusRequired = True
end if

gi_Mode = c_MODE_ACCESS
Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select


Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "Select * FROM Z_Floors Where JOB = '" & JOB & "'"
		rs.Cursortype = GetDBCursorTypeInsert
		rs.Locktype = GetDBLockTypeInsert
		rs.Open strSQL, DBConnection
		
		if JobStatusRequired = False then
			if JOB <> "" and FLoor <> "" then
				rs.filter = "JOB = '" & JOB & "' and Floor = '" & Floor & "' and ParentFloor = '" & ParentFloor & "'"	
				
				if rs.eof then 
					rs.AddNew
					rs.Fields("Job") = JOB
					rs.Fields("Floor") = Floor
					rs.Fields("ParentFloor") = ParentFloor
					rs.Fields("JobStatus") = JobStatus

					If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
					rs.update
					Call StoreID1(isSQLServer, rs.Fields("ID"))
				else 
					AlreadyEntered = TRUE
				end if
			end if
		end if
End Function

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "Select JOB,Completed FROM Z_Jobs Where Completed =FALSE ORDER BY JOB ASC"
	rs2.Cursortype = GetDBCursorTypeInsert
	rs2.Locktype = GetDBLockTypeInsert
	rs2.Open strSQL2, DBConnection
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	
	if Job = "" then
		strSQL3 = "Select * FROM Z_FLoors ORDER BY Floor ASC"
	else
		strSQL3 = "Select * FROM Z_FLoors WHERE JOB = '" & JOB & "' ORDER BY Floor ASC"
	end if
	
	rs3.Cursortype = GetDBCursorTypeInsert
	rs3.Locktype = GetDBLockTypeInsert
	rs3.Open strSQL3, DBConnection


%>


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job / Colour</a>
        </div>

              <form id="Floors" title="Floors" class="panel" name="Floors" action="AllFloorsReport.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Floor </h2>

<fieldset>

    <div class="row">
        <label> Job </label>
        <select name ='JOB'>
		
		<%
		if len(JOB) = 3 then
			response.write "<option value = '"
			response.write JOB
			response.write "'>"
			response.write JOB
			response.write "</option> "
		end if
	
		Do While Not rs2.eof
			response.write "<option value = '"
			response.write rs2("JOB")
			response.write "'>"
			response.write rs2("JOB")
			response.write "</option> "
		rs2.movenext 
		loop
		
		%>
		</select>
    </div>

    <div class="row">
        <label> Floor </label>
        <input type="text" name='Floor' id='Floor' >
    </div>
	
	<div class="row">
        <label> Parent Floor </label>
			<input type="text" name='ParentFloor' id='ParentFloor' >
    </div>
	
	<div class="row">
	<label> Job Status</label>
	<select name ='JOBSTATUS' required>
				<option value = 'Guaranteed'>Guaranteed</option>
				<option value = 'Measured'>Measured</option>
			</select>
	</div>
	
</fieldset>
	    <a class="whiteButton" href="javascript:Floors.submit()">Submit</a>

      

<ul id="Profiles" title="Glass" selected="true">



<li class = 'group'>Floor Details <%if JOB <> "" then response.write " - " & JOB end if %></li>

<%
if AlreadyEntered = TRUE then
	response.write "<li>No New Entry, Floor Already Exists</li>"
end if
if JobStatusRequired = TRUE then
	response.write "<li>Must Choose Guaranteed / Measured / Mixed - Cannot leave blank</li>"
end if

%>

<li>
	<table border = "1">
	<TR><TH>Job</TH><TH>Floor</TH><TH>Parent</TH><TH>Status</TH></TR>
<%

rs3.movefirst
do while not rs3.eof
	response.write "<TR>"
	response.write "<TD> " & rs3("Job") & " </TD> "
	response.write "<TD> " & rs3("Floor") & " </TD> "
	response.write "<TD> " & rs3("ParentFloor") & " </TD> "
	response.write "<TD> " & rs3("JobStatus") & " </TD> "
	response.write "</TR>"
rs3.movenext
loop
	
%>
	</ul>      
            </form>
                
<%    


      
DbCloseAll
     
%>
</body>
</html>
