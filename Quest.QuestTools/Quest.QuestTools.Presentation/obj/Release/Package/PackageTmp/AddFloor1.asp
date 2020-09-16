<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
	 "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<!--#include file="dbpath.asp"-->

	<!-- Add Floor Function adds Floor information to the Z_Floors Table -->
	<!-- First Checks Job Status of Guaranteed / Measured / Mixed -- if Mixed, forces choice by floor -->
	<!-- Continues to allow new floors to be entered as Parent or Child -->
	<!-- Michael Bernholtz, October 2018 - for Ariel Aziza, David Ofir and Jody Cash -->
				
	<!-- Add Floor Page 1 of 3 -->
		<!-- Enter Job and send to Page 2, Page 2 checks Job table for JobStatus-->	
			

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
        <h1 id="pageTitle">Add Floors</h1>
                <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self" >Jobs</a>
    </div>
    
              <form id="add" title="Enter Floors" class="panel" name="add" action="addFloor2.asp" method="GET" target="_self" selected="true" > 
        <h2>Select Job</h2>
  
   

<fieldset>
     <div class="row">
<label> Job </label>
        <select name ='JOB'>
<%	


Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "Select JOB,Completed FROM Z_Jobs Where Completed =FALSE ORDER BY JOB ASC"
	rs2.Cursortype = GetDBCursorTypeInsert
	rs2.Locktype = GetDBLockTypeInsert
	rs2.Open strSQL2, DBConnection
	
		Do While Not rs2.eof
			response.write "<option value = '"
			response.write rs2("JOB")
			response.write "'>"
			response.write rs2("JOB")
			response.write "</option> "
		rs2.movenext 
		loop
		
		rs2.close
		set rs2 = nothing

		%>
		</select>
            </div>
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:add.submit()">Continue to Add Floors</a><BR>
            
            
            </form> 

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

