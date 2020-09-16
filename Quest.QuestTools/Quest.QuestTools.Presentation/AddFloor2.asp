<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
	<!-- Add Floor Function adds Floor information to the Z_Floors Table -->
	<!-- First Checks Job Status of Guaranteed / Measured / Mixed -- if Mixed, forces choice by floor -->
	<!-- Continues to allow new floors to be entered as Parent or Child -->
	<!-- Michael Bernholtz, October 2018 - for Ariel Aziza, David Ofir and Jody Cash -->
				
	<!-- Add Floor Page 2 of 3 -->
		<!-- Checks Job table for JobStatus Mixed Gets choice, Guaranteed or Measured stay Greyed-->	
		<!-- Add Floor information -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Add Floors</title>
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




    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="AddFloor1.asp" target="_self" >Choose Job</a>
    </div>
    
 <% 
JOB = REQUEST.QueryString("JOB")


	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_Jobs Where JOB = '" & JOB & "' ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Z_Floors Where JOB = '" & JOB & "' ORDER BY Floor ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT Distinct ParentFloor FROM Z_Floors Where JOB = '" & JOB & "' ORDER BY ParentFloor ASC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

JOBSTATUS = rs("JobStatus")
Select Case JobStatus
	Case "Guaranteed", "Measured"
		EntryForm = "Simple"
	Case "Mixed"
		EntryForm = "Mixed"
	Case Else
		EntryForm = "Error"
End Select

%>
	       
    
	<form id="add" title="Add Floors" class="panel" name="add" action="AddFloor2Conf.asp" method="GET" target="_self" selected="true" > 
		<h2>Job Status of <%response.write Job%> is <%response.write JOBSTATUS%></h2>

		<fieldset>
		
		<input type="hidden" name='Job' id='Job' value ='<% response.write JOB%>' >
			 <div class="row">
				<label>Floor</label>
				<input type="text" name='Floor' id='Floor'>
			</div>
			 <div class="row">
				<label>Parent</Label>
				<input type="checkbox" name="Parent" value="Parent" /></TD>
			</div>
			<div class="row">
				<label>Parent (if Different)</label>
				 <select name ='ParentFloor'>
			<%	
				Do While Not rs3.eof
					response.write "<option value = '"
					response.write rs3("ParentFloor")
					response.write "'>"
					response.write rs3("ParentFloor")
					response.write "</option> "
				rs3.movenext 
				loop
				rs3.close
				set rs3 = nothing
			%>
			</select>
			</div>
				
<%
	if EntryForm = "Simple" then
%>
	<input type="hidden" name='JobStatus' id='JobStatus' value ='<% response.write JOBSTATUS%>' >
<%
	end if 
%>

				
		</fieldset>


        <BR>
	
	<%
	if EntryForm = "Error" then
	%>
		<p> Job Status not Created in Job Table, This Page cannot Submit until this is updated</p>
	<%
	Else
	%>
        <a class="whiteButton" href="javascript:add.submit()">Submit Floor</a><BR>
     <%
	End if
	%>	    
	
	
	<ul id="Profiles" title="Glass" selected="true">



<li class = 'group'>Floor Details <%if JOB <> "" then response.write " - " & JOB end if %></li>

<li>
	<table border = "1">
	<TR><TH>Job</TH><TH>Floor</TH><TH>Parent</TH><TH>Status</TH></TR>
<%
if not rs2.eof then
	rs2.movefirst
end if
do while not rs2.eof
	response.write "<TR>"
	response.write "<TD> " & rs2("Job") & " </TD> "
	response.write "<TD> " & rs2("Floor") & " </TD> "
	response.write "<TD> " & rs2("ParentFloor") & " </TD> "
	if rs2("jobStatus") <> "Guaranteed" and rs2("jobStatus") <> "Measured" then
		response.write "<TD>Required</TD> "
	else
		response.write "<TD> " & rs2("JobStatus") & " </TD> "
	end if
	response.write "</TR>"
rs2.movenext
loop
rs2.close
set rs2 = nothing
	
%>
</table>
</li>
	
	
<%
	if EntryForm = "Mixed" then
%>	
			<input type="submit" value = "Select Job Status" class="greenButton" onclick="add.action='addfloor3.asp';"></input>
<%
	end if
%>
</ul> 
            </form>
            
 
            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

