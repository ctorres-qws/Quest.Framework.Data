<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search Wareghouses by PO/Bundle Search page -->
<!--Created May 6th, by Michael Bernholtz -Overarching tool -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Warehouses</title>
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
filterpo = request.QueryString("PO")
if filterpo = "" then
	filterpo = request.QueryString("POSEARCH")
end if
warehouse = request.QueryString("warehouse")
if warehouse = "" then
	warehouse = "GOREWAY"
end if

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE='" & warehouse & "' AND (PO LIKE '%" & filterpo & "%' OR BUNDLE LIKE '%" & filterpo & "%' OR EXBUNDLE LIKE '%" & filterpo & "%' OR COLORPO LIKE '%" & filterpo & "%') ORDER BY WAREHOUSE, PART"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="warehousebypo1.asp" target="_self">Warehouse/PO</a>
        </div>

        <ul id="Profiles" title="Warehouse Stock" selected="true">

<%

response.write "<li class='group'>" & warehouse & " STOCK</li>"

do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
colour = rs("colour")
PO = rs("po")
Bundle= rs("Bundle")
EXBundle= rs("ExBundle")
ColorPO = rs("ColorPO")

aisle = rs("aisle")
rack = rs("Rack")
%>
<!--response.write part & ", " & qty & " SL"  & ", " & Colour -->
<li><a href="stockbyrackedit.asp?ticket=warehouse&PO=<%response.write filterpo %>&id=<% response.write id %>" target="_self"> 
<%response.write part & ", "%>
<%response.write Colour & ": "%>
<%response.write  qty & " SL " %>
<%response.write "PO: " & PO & " "%>
<%response.write "Bundle: " & Bundle & " "%>
<%response.write "ExBundle: " & ExBundle & " "%>
<%response.write "ColorPO: " & ColorPO & " "%>
</a></li>
<%

rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
