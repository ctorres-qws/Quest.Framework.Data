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
<%
JOB = Request.QueryString("Job")
%>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
               <a class="button leftButton" type="cancel" href="AddFloor2.asp?Job=<%response.write JOB%>" target="_self" >Enter Floor</a>
    </div>
    
  <%
  
  
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_Floors Where JOB = '" & JOB & "' ORDER BY Floor ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection  
	
%>   
    
              <form id="edit" title="Select Job Status" class="panel" name="edit" action="addFloor3Conf.asp" method="GET" target="_self" selected="true" >
        <h2>Select Job Status</h2>
  
   

<fieldset>
<input type="hidden" name='Job' id='Job' value ='<% response.write JOB%>' >
     <div class="row">
                
			<table border = "1">
	<TR><TH>Job</TH><TH>Floor</TH><TH>Parent</TH><TH>Status</TH></TR>
<%



do while not rs.eof
	
	
	
		response.write "<TR>"
		response.write "<TD> " & rs("Job") & " </TD> "
		response.write "<TD> " & rs("Floor") & " </TD> "
		response.write "<TD> " & rs("ParentFloor") & " </TD> "
		
		if rs("JobStatus") = "Guaranteed" or rs("JobStatus") = "Measured" then
			response.write "<TD> " & rs("JobStatus") & " </TD> "
		else
		
			response.write "<TD>"
			%>
			<select name ='<%response.write rs("ID")%>' required>
						<option value = '-'>-</option>
						<option value = 'Guaranteed'>Guaranteed</option>
						<option value = 'Measured'>Measured</option>
				</select>

			<%
			Response.write "</TD>"

		
		end if
		response.write "</TR>"
rs.movenext
loop
rs.close
set rs = nothing
	
%>
</table>	
				
            </div>
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Enter Status</a><BR>
                       
            </form>
            
    
</body>
</html>

<% 

DBConnection.close
set DBConnection=nothing
%>

