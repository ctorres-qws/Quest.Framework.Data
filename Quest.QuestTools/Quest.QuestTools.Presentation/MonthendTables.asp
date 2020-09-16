                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Creates 2 tables to be used for Month end calculations-->
<!-- Last 4 Days of Inventory in Window Production and Last 4 Days of Trucks-->
<!-- Y_INV and X_SHIPPING_TRUCK-->
<!-- ALSO uses Y_Master for Pricing-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Month End 4 Day</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
	iui.animOn = true;
	
	function exportTableToExcel(tableID){
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    var tableSelect = document.getElementById(tableID);
    var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
    
    // Specify file name
    var filename = 'excel_data.xls';
    
    // Create download link element
    downloadLink = document.createElement("a");
    
    window.open('data:' + dataType + ', ' + tableHTML);
    
}
    </script>
    
    <%
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT top 2000 * FROM Y_INV  Where WAREHOUSE = 'WINDOW PRODUCTION' Order BY ID DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT top 2000 * FROM Y_INV  Where WAREHOUSE = 'JUPITER PRODUCTION' Order BY ID DESC"
rs4.Cursortype = GetDBCursorType
rs4.Locktype = GetDBLockType
rs4.Open strSQL4, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER"
'rs2.Cursortype = GetDBCursorType
'rs2.Locktype = GetDBLockType
'rs2.Open strSQL2, DBConnection
Set rs2 = GetDisconnectedRS(strSQL2, DBConnection)

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT top 25 * FROM X_SHIP_TRUCK Order BY ID DESC"
rs3.Cursortype = GetDBCursorType
rs3.Locktype = GetDBLockType
rs3.Open strSQL3, DBConnection

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT top 25 * FROM [MS Access;DATABASE=\\10.34.16.11\db\Scan_Texas_Dev.mdb].X_SHIPPING_TRUCK Order BY ID DESC"
rs5.Cursortype = GetDBCursorType
rs5.Locktype = GetDBLockType
rs5.Open strSQL5, DBConnection


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Window Production Inventory and Trucks" selected="true">
        
        
<% 
response.write "<li class='group'>Window Production </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li>WINDOW PRODUCTION <button id=""btnWindowProduction"" onclick=""exportTableToExcel('windowProduction')"" style='cursor:pointer' class=""button style='top: -50px;' rightButton"" title=""Download Report"" onclick=""exportTableToExcel()"">Download</button></li>  "
response.write "<li><table border='1' id='windowProduction' class='sortable'><tr><th>Part</th><th>Qty</th><th>Length(Feet)</th><th>Colour</th><th>Dateout</th><th>Floor</th><th>Price(KGM)</th><th>Price (LBF)</th><th>Inventory Type</th><th>Description </th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Part") & "</td>"
	response.write "<td>" & RS("Qty") & "</td>"
	response.write "<td>" & RS("LFT") & "</td>"
	response.write "<td>" & RS("Colour") & "</td>"
	response.write "<td>" & RS("DateOut") & "</td>"
	response.write "<td>" & RS("Note") & "</td>"
	rs2.filter = "Part = '" & RS("Part") & "'"
	if rs2.eof then
	response.write "<td>N/A</td><td></td><td>N/A</td><td></td>"
	else
		response.write "<td>" & RS2("KGM") & "</td>"
		response.write "<td>" & RS2("LBF") & "</td>"
		response.write "<td>" & RS2("InventoryType") & "</td>"
		response.write "<td>" & RS2("Description") & "</td>"
	end if 
	rs2.filter = ""
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing

response.write "<li> JUPITER PRODUCTION</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Part</th><th>Qty</th><th>Length(Feet)</th><th>Colour</th><th>Dateout</th><th>Floor</th><th>Price(KGM)</th><th>Price (LBF)</th><th>Inventory Type</th><th>Description </th></tr>"
do while not rs4.eof
	response.write "<tr><td>" & rs4("Part") & "</td>"
	response.write "<td>" & rs4("Qty") & "</td>"
	response.write "<td>" & rs4("LFT") & "</td>"
	response.write "<td>" & rs4("Colour") & "</td>"
	response.write "<td>" & rs4("DateOut") & "</td>"
	response.write "<td>" & rs4("Note") & "</td>"
	rs2.filter = "Part = '" & rs4("Part") & "'"
	if rs2.eof then
	response.write "<td>N/A</td><td></td><td>N/A</td><td></td>"
	else
		response.write "<td>" & rs2("KGM") & "</td>"
		response.write "<td>" & rs2("LBF") & "</td>"
		response.write "<td>" & rs2("InventoryType") & "</td>"
		response.write "<td>" & rs2("Description") & "</td>"
	end if 
	rs2.filter = ""
	response.write " </tr>"
	rs4.movenext
loop
response.write "</table></li>"


rs4.close
set rs4 = nothing
rs2.close
set rs2 = nothing

response.write "<li class='group'>Trucks CANADA </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>TruckName</th><th>ID</th><th>Job/Floor</th><th>Shipdate</th><th>Open Date</th><th>Active</th></tr>"
do while not rs3.eof
	response.write "<tr><td>" & RS3("truckname") & "</td>"
	response.write "<td>" & RS3("ID") & "</td>"
	'response.write "<td>" & RS3("job") & "</td>"
	'response.write "<td>" & RS3("floor") & "</td>"
	response.write "<td style='word-break:break-all;'>" & RS3("slist") & "</td>"
	response.write "<td>" & RS3("shipdate") & "</td>"
	response.write "<td>" & RS3("createdate") & "</td>"
	response.write "<td>" & RS3("Active") & "</td>"
	response.write " </tr>"
	rs3.movenext
loop
response.write "</table></li>"



rs3.close
set rs3 = nothing

response.write "<li class='group'>Trucks USA</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>TruckName</th><th>ID</th><th>Job</th><th>Floor</th><th>Shipdate</th><th>Open Date</th><th>Active</th></tr>"
do while not rs5.eof
	response.write "<tr><td>" & rs5("truckname") & "</td>"
	response.write "<td>" & rs5("ID") & "</td>"
	response.write "<td>" & rs5("job") & "</td>"
	response.write "<td>" & rs5("floor") & "</td>"
	response.write "<td>" & rs5("shipdate") & "</td>"
	response.write "<td>" & rs5("createdate") & "</td>"
	response.write "<td>" & rs5("Active") & "</td>"
	response.write " </tr>"
	rs5.movenext
loop
response.write "</table></li></ul>"



rs5.close
set rs5 = nothing


DBConnection.close 
set DBConnection = nothing
%>
</body>
</html>
