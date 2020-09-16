                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title> TESTING</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    
    <%
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE ( DATEOUT = #01/26/2017# or  DATEOUT =#01/27/2017# or DATEOUT = #01/29/2017# or DATEOUT = #01/30/2017# or DATEOUT = #01/31/2017# ) AND WAREHOUSE = 'WINDOW PRODUCTION' ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT PART, KGM, LBF FROM Y_MASTER"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Service" selected="true">
        
        
<% 
response.write "<li class='group'>SERVICE GLASS REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Part</th><th>Floor</th><th>QTY</th><th>Length in Feet</th><th>Job/Colour</th>><th>WIP DATE</th><th>KGM</th><th>LBF</th><th>QTY * Length * Price</th></tr>"
do while not rs.eof

	response.write "<tr>"
	response.write "<td>" & RS("PART") & "</td>"
	response.write "<td>" & RS("NOTE") & "</td>"
	response.write "<td>" & RS("QTY") & "</td>"
	response.write "<td>" & RS("LFT") & "</td>"
	response.write "<td>" & RS("COLOUR") & "</td>"
	response.write "<td>" & RS("DateOUT") & "</td>"
	
	rs2.filter = ""
	rs2.filter = " PART = '" & rs("PART") & "'"
	if rs2.eof then
	else
	response.write "<td>" & RS2("KGM") & "</td>"
	response.write "<td>" & RS2("LBF") & "</td>"
	
		if RS2("KGM") > 0 then 
			response.write "<td>" & RS("QTY")*RS("LFT")*RS2("KGM") & "</td>"
		end if
		if RS2("LBF") > 0 then 
			response.write "<td>" & RS("QTY")*RS("LFT")*RS2("LBF") & "</td>"
		end if
	end if
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li></ul>"



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
%>
</body>
</html>
