<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

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
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' or WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'NPREP' ORDER BY WAREHOUSE, PART"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)
%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="stockgbypo1.asp" target="_self">Stock by PO</a>
        </div>

        <ul id="Profiles" title="Profiles" selected="true">

<% 
filterpo = request.QueryString("PO")
if filterpo = "" then
	filterpo = request.QueryString("POSEARCH")
end if
rs.filter = "WAREHOUSE='GOREWAY' AND PO ='" & filterpo & "'"
'"DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear

response.write "<li class='group'>GOREWAY STOCK</li>"

do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
colour = rs("colour")
PO = rs("po")

%>

<li><a href="stockbyrackedit.asp?ticket=goreway&PO=<%response.write filterpo %>&id=<% response.write id %>" target="_self"> <%response.write part & ", " & qty & " SL"  & ", " & Colour %></a></li>
<%

aisle = rs("aisle")
rack = rs("Rack")
rs.movenext
loop

rs.filter = "WAREHOUSE='NASHUA' AND PO ='" & filterpo & "'"
'"DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear

response.write "<li class='group'>NASHUA STOCK</li>"

do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
colour = rs("colour")
PO = rs("po")

%>

<li><a href="stockbyrackedit.asp?ticket=goreway&PO=<%response.write filterpo %>&id=<% response.write id %>" target="_self"> <%response.write part & ", " & qty & " SL"  & ", " & Colour %></a></li>
<%

aisle = rs("aisle")
rack = rs("Rack")
rs.movenext
loop

rs.filter = "WAREHOUSE='NPREP' AND PO ='" & filterpo & "'"
'"DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear

response.write "<li class='group'>NASHUA PREP STOCK</li>"

do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
colour = rs("colour")
PO = rs("po")

%>

<li><a href="stockbyrackedit.asp?ticket=goreway&PO=<%response.write filterpo %>&id=<% response.write id %>" target="_self"> <%response.write part & ", " & qty & " SL"  & ", " & Colour %></a></li>
<%

aisle = rs("aisle")
rack = rs("Rack")
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
