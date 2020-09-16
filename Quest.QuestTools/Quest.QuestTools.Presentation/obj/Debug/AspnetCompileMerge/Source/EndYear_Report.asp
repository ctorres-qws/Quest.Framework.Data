<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
			<!-- Designed December 2019 as a Table location for end of year Windows completed but not shipped. -->
			<!-- Scan_Endyear.asp is scanner-->
			<!-- X_SHIP_ENDYEAR is Database table-->
			<!-- EndYear_Report.asp is the Report to view it-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>End Year Report</title>
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
strSQL = "SELECT * FROM X_SHIP_ENDYEAR ORDER BY Barcode ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Counter = 0
CurrentJob = "XXXX"

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">END YEAR</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Shipping</a>
        </div>

        <ul id="Profiles" title="Windows Completed" selected="true">
<% 
response.write "<li class='group'>Windows Completed, but not Shipped</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "

do while not rs.eof
Counter = Counter+1
LastJob = CurrentJob
CurrentJob = Left(RS("Barcode"),3)

if LastJob = CurrentJob  then
else
	if LastJob = "XXXX" then
	else
		Response.write "<tr><td><B>Counter: " & Counter & "</B></td><td></td></tr>"
		Response.write "</table></li>"
	end if
	Counter = 0
	response.write "<li><table border='1' class='sortable'><tr><th>Barcode</th><th>Scan Date</th></tr>"
	
end if


		response.write "<tr>"
		response.write "<td>" & RS("Barcode") & "</td><td>" & RS("SCANDATE") &"</td>"
		response.write " </tr>"
	rs.movenext

loop
Response.write "<tr><td><B>Counter: " & Counter+1 & "</B></td><td></td></tr>"
Response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
      </ul>
</body>
</html>
