<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search SAPA Stock by PO Search page -->
<!--Created May 6th, by Michael Bernholtz at Request of Ruslan Bedoev/Ariel Azia -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search SAPA Inventory </title>
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE='SAPA' AND PO ='" & filterpo & "' ORDER BY WAREHOUSE, PART"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="sapabypo1.asp" target="_self">SAPA by PO</a>
        </div>

        <ul id="Profiles" title="SAPA Stock" selected="true">

<% 

'rs.filter = "WAREHOUSE='SAPA' AND PO ='" & filterpo & "'"
'"DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear

response.write "<li class='group'>SAPA STOCK</li>"

do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
colour = rs("colour")
PO = rs("po")
bundle = rs("bundle")

aisle = rs("aisle")
rack = rs("Rack")
%>

<li><a href="stockbyrackedit.asp?ticket=sapa&PO=<%response.write filterpo %>&id=<% response.write id %>" target="_self"> <%response.write part & ", " & qty & " SL"  & ", " & Colour & " / " & bundle %></a><img src='/partpic/<%response.write part%>.png'/></li>"</li>
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
