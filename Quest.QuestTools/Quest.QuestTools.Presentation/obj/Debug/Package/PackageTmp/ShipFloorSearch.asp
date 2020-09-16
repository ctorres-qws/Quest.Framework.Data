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

	    
    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle">Old Trucks</h1>
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
   
 
        <form id="ShipSearch" title="Shipping Search" class="panel" name="ShipSearch" action="ShipFloorViewer.asp" method="GET" selected="true">
        <h2>Choose Job and Floor to look for in Old Truck list</h2>
        <fieldset>
       

		<div class="row">    
			<label>JOB</label>
			<select name='Job' id='Job' class ='leftinput'">
<%

				Set rs = Server.CreateObject("adodb.recordset")
				strSQL = "SELECT JOB FROM Z_JOBS WHERE COMPLETED = FALSE ORDER BY JOB ASC"
				rs.Cursortype = 2
				rs.Locktype = 3
				rs.Open strSQL, DBConnection

				rs.movefirst
				Do While Not rs.eof

				Response.Write "<option value='"
				Response.Write rs("JOB")
				Response.Write "'>"
				Response.Write rs("JOB")
				Response.write "</option>"

				rs.movenext

				loop
				rs.close
				set rs=nothing
%>
			</select>
		</div>
		
		<div class="row">
			<label>Floor</label>
			<input type="text" name='Floor' id='Floor'  required>
			<input type="hidden" name='ticket' id='ticket' value='search'>
		</div>
		
		
		<div class="row">
		<a class="whiteButton" onClick="ShipSearch.submit()">Search Floor</a><BR>
		</div>
        </fieldset>	
	</form> 


<%
rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing
%>    
  
</body>
</html>
