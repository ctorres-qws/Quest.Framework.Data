<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Ruslan Inventory list with edit ability in order to quickly update old or outdated inventory items to be removed or moved from warehouse.-->
<!--Created by Michael Bernholtz at Request of Ruslan Bedoav March 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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
    
    <%
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE='GOREWAY' ORDER BY AISLE, RACK, SHELF DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
'Set rs = GetDisconnectedRS(strSQL, DBConnection)
%>
 
    </head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.html#_Inv" target="_self">Stock</a>
        </div>
        <ul id="Profiles" title="Profiles" selected="true">
<%

'create a new button, with checkboxes
'Add a column to Y_INV
'Checkbox sets that to one 
'Logic runs on all where checkbox != 1
' Drop added column

'rs.filter = "WAREHOUSE='GOREWAY'"

response.write "<li>All Items in Inventory</li>"
response.write "<li><table border='1' class='sortable' width='75%'><tr><th>Aisle</th><th>part</th><th>PO</th><th>Bundle</th><th>Colour</th><th>QTY</th><th>Button</th></tr>"
int counter

	rs.movefirst
	do while not rs.eof
		aisle = trim(rs("aisle"))
		part = trim(rs("part"))
		qty = trim(rs("qty"))
		po = trim(rs("PO"))
		bundle = trim(rs("bundle"))
		colour = trim(rs("colour"))


		response.write "<tr><td>" & aisle & "</td><td>" & part & "</td><td>" & PO & "</td><td>" & bundle & "</td><td>" & Colour & "</td><td>" & qty & "</td>"  '<td><input type='checkbox' id='exists name='" & trim(rs("id")) & "'</td>"	
		response.write "<td><a class='whiteButton' target='#_self' href='stockbyrackedit.asp?id=" & trim(rs("id")) & "&aisle=" & trim(rs("aisle")) & "'>Edit</a></td>"
	rs.movenext
	loop
	response.write "</table></li>"
	rs.close
	Set rs = nothing
	DBConnection.close
	Set DBConnection = nothing
	
%>

      </ul>
</body>
</html>
