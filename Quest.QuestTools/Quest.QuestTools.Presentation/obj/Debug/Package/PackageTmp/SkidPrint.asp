<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Print - Views one Skid of choice and then prints -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Skid Views</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	
	<script>
	function printbutton(){
			window.print()
}
</script>
	
	
<%

Skidname = REQUEST.QueryString("skids")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM SKIDITEM ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "name = '" & skidname & "'"

%> 
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Skid REPORT</h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
        <ul id="Profiles" title="Skids and Skid Items" selected="true">

<%


	response.write " <li>Contents of  " & Skidname 
	response.write "<input type ='submit' value = 'Print' onclick='printbutton()'></li>"
	response.write "<li><table border='1' class='sortable' width='75%'><tr><th width='30%'>Barcode</th><th width='10%'>Job</th><th width='10%'>Floor</th><th width='10%'>Tag</th><th  width='20%'>Scan Date</th><th  width='20%'>Flushed</th></tr>"
	
	if not rs.bof then
		rs.movefirst
		Do while not rs.eof

			response.write "<tr><td>" & rs("Barcode") & "</td><td>" & rs("Job") & "</td><td>" & rs("Floor") & "</td><td>" & rs("Tag") & "</td><td>" & rs("ScanDate") & "</td><td>" & rs("FlushedDate") & "</td></tr>"
			
		rs.movenext
		loop
		response.write "</table></li>"
	
	else
		Response.write "<tr> <td> Empty Skid - Please try again </td></tr>"
	
	end if

rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing

%>


</body>
</html>	