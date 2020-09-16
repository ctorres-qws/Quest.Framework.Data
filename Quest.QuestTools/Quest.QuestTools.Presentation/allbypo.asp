<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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

filterpo = request.QueryString("po")
if filterpo = "" then
	filterpo = request.QueryString("POSEARCH")
end if

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE PO Like '%" & filterpo & "%' ORDER BY WAREHOUSE, PART"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="allbypo1.asp" target="_self">PO</a>
        </div>

        <ul id="Profiles" title="Profiles" selected="true">
<% 

'Loops through all Warehouses and then posts for each one.

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE ORDER BY ID ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

rs2.movefirst
Do While Not rs2.eof
	WarehouseName = rs2("Name")

	' Each Warehouse
	rs.filter = "Warehouse = '" & WarehouseName & "'" 

	response.write "<li class='group'>" & WarehouseName & "</li>"
	do while not rs.eof
%>
		<li><a href="stockbyrackedit.asp?ticket=allpo&PO=<%response.write filterpo %>&id=<% response.write rs("ID") %>" target="_self"> <%response.write rs("part") & ", " & rs("qty") & " SL"  & ", " & rs("colour") & " - " & rs("bundle")  %></a></li>
<%

	rs.movenext
	loop

rs2.movenext
loop

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
set DBConnection = nothing

%>
      </ul>

</body>
</html>
