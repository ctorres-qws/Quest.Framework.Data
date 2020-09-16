<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search Production Stock by PO Search page -->
<!--Created May 1st, by Michael Bernholtz at Request of Ruslan Bedoev -->

<!-- February 2019 - Added USA view -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Production Inventory </title>
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


if CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE='JUPITER PRODUCTION' ) AND (PO LIKE '%" & filterpo & "%' OR BUNDLE LIKE '%" & filterpo & "%' OR EXBUNDLE LIKE '%" & filterpo & "%' OR COLORPO LIKE '%" & filterpo & "%') ORDER BY WAREHOUSE, PART"
else
	strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE='WINDOW PRODUCTION' OR WAREHOUSE = 'COM PRODUCTION') AND (PO LIKE '%" & filterpo & "%' OR BUNDLE LIKE '%" & filterpo & "%' OR EXBUNDLE LIKE '%" & filterpo & "%' OR COLORPO LIKE '%" & filterpo & "%')) ORDER BY WAREHOUSE, PART"
end if

Set rs = Server.CreateObject("adodb.recordset")

'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)
%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="productionbypo1.asp" target="_self">Production by PO/Bundle</a>
        </div>

        <ul id="Profiles" title="Production Stock" selected="true">

<%

Response.Write "<li class='group'>Production STOCK</li>"

Do While Not rs.eof
	part = rs("part")
	qty = rs("qty")
	id = rs("ID")
	shelf = rs("shelf")
	colour = rs("colour")
	PO = rs("po")
	bundle = rs("bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=prod&PO=<%Response.Write filterpo %>&id=<% Response.Write id %>" target="_self"> <%Response.Write part & ", " & qty & " SL"  & ", " & Colour & " / " & bundle %></a></li>
<%

	aisle = rs("aisle")
	rack = rs("Rack")
	rs.movenext
Loop

rs.close
set rs = nothing
DBConnection.close
set DBConnection=nothing

%>
      </ul>

</body>
</html>
