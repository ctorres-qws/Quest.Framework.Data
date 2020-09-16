<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 <!--MILVAN version of Stock levels, created for Kevin Cosgrove by Michael Bernholtz, Jan 2020 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels - MILVAN</title>
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
  
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="StockLevelsDH.asp" target="_self">Combined View</a>
    </div>

<ul id="screen1" title="Stock Level - MILVAN" selected="true">            

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'MILVAN' order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - MILVAN Only</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Unallocated Mill</th><th>Allocated Mill</th><th>Painted Stock</th><th>Total Qty</th></tr>"

rs2.movefirst
do while not rs2.eof
	partqty = 0
	allocatedqty = 0
	Totalqty = 0
	paintedqty = 0

	rs.movefirst
	do while not rs.eof
		IF rs2("Part") = rs("part") then

		Totalqty = rs("Qty") + Totalqty
			if rs("colour") = "Mill" then
				if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
					partqty = rs("Qty") + partqty
				else
					allocatedqty = rs("Qty") + allocatedqty
				End if
			else
				paintedqty = rs("Qty") + paintedqty
			End if
		End if
	rs.movenext
	loop

	if partqty = 0 and allocatedqty = 0 and paintedqty = 0 then
	else
		response.write "<tr><td><a href='stockbydieMILVAN.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td><td>" & rs2("Description") & "</td><td>" & partqty & "</td><td>" & allocatedqty & "</td><td>" & paintedqty & "</td><td>" & Totalqty & "</td></tr>"
	end if 

rs2.movenext
loop
response.write "</table></li>"

%>

   </ul>
</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

