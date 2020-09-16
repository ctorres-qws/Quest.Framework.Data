<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<%
ScanMode = True
%>
<!--#include file="Texas_dbpath.asp"-->
<!--Shipping Table Deleted Reporting - July 2019, Michael Bernholtz -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Removed Shipping items</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	    
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
 <div class="toolbar">
        <h1 id="pageTitle">Removed Windows</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "ShipHomeManager.HTML"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "ShipHomeManager.HTML"
				HomeSiteSuffix = ""
			end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
   
 
        <ul id="Profiles" title="Deleted Scans" selected="true">
		
		<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM X_SHIP Where [Deleted] = TRUE ORDER BY ID ASC")
rs.Cursortype = 1
rs.Locktype = 3
if CountryLocation = "USA" then
	rs.Open strSQL, DBConnection_Texas
else	
	rs.Open strSQL, DBConnection
end if
RemovedCount = Rs.RecordCount

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = FixSQL("SELECT * FROM X_SHIP Where [Deleted] = FALSE ORDER BY ID ASC")
rs2.Cursortype = 1
rs2.Locktype = 3
if CountryLocation = "USA" then
	rs2.Open strSQL2, DBConnection_Texas
else	
	rs2.Open strSQL2, DBConnection
end if

	%>


<Table><TR><TH>All items Unscanned NOT scanned back after</TH><TH>All Unscans (<% Response.write RemovedCount %>)</TH></TR>
<TR><TD>
<table border='1' class='sortable'><tr><th>Truck</th><th>Barcode</th><th>ShipDate</th><th>Deleted Date</th></tr>
<%
rs.filter =""
Do While Not rs.eof

		BarcodeTest = RS("Barcode")
		rs2.filter = " Barcode = '" & BarcodeTest & "'"
		if rs2.eof then
			response.write "<tr>"
			response.write "<td>" & RS("Truck") & "</td>"
			response.write "<td>" & RS("Barcode") & "</td>"
			response.write "<td>" & RS("ShipDate") & " " & RS("ShipTime") & "</td>"
			response.write "<td>" & RS("DeleteDate") & "</td>"
			response.write "</tr>"
	end if

rs.movenext
loop

%>
</table>
</TD><TD>
<table border='1' class='sortable'><tr><th>Truck</th><th>Barcode</th><th>ShipDate</th><th>Deleted Date</th></tr>
<%
rs.filter =""
Do While Not rs.eof

	response.write "<tr>"
	response.write "<td>" & RS("Truck") & "</td>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("ShipDate") & " " & RS("ShipTime") & "</td>"
	response.write "<td>" & RS("DeleteDate") & "</td>"
	response.write "</tr>"

	
rs.movenext
loop

%>
</table>
</TD></TR>
</Table></li>



<%
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing
DBConnection_Texas.close 
set DBConnection_Texas = nothing


%>
               
    </ul>      
  
</body>
</html>
