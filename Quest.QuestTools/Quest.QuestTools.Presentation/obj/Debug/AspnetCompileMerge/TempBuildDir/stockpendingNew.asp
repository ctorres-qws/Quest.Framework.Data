<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stockpending.asp updated as a new Table form page was created stockpendingtable.asp, May 23rd, 2014-->
<!-- Updated December 2014 - Added Metra -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Pending</title>
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
'b_SQL_Server = true
'Set DBConnection = Server.CreateObject("adodb.connection")
'DSN = GetConnectionStr(true) 'method in @common.asp
'DBConnection.Open DSN

%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
        <ul id="Profiles" title="Pending Stock" selected="true">

         <li class="group"><a href="stockpendingtable.asp?part=<%response.write part%>" target="_self" >Stock Pending (Row Form) - Switch to Table Form</a></li>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV YI INNER JOIN Y_MASTER YM ON YM.Part = YI.Part WHERE YI.WAREHOUSE IN('SAPA','HYDRO','DURAPAINT(WIP)','DEPENDABLE','EXTAL SEA','KEYMARK','TILTON(WIP)','CAN-ART','APEL','METRA') "
If b_SQL_Server Then
	strSQL = strSQL & " ORDER BY Case WHEN YI.WareHouse = 'SAPA' or YI.WareHouse = 'HYDRO' Then 2 When YI.WareHouse = 'DURAPAINT(WIP)' THEN 1 ELSE 0 END DESC, YI.Part ASC"
Else
	strSQL = strSQL & " ORDER BY IIF(YI.WAREHOUSE='SAPA' or YI.WareHouse = 'HYDRO' ,2,IIF(YI.WAREHOUSE='DURAPAINT(WIP)',1,0)) DESC, YI.PART ASC"
End If

rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

warehouse_prev= ""
do while not rs.eof
	warehouse = rs("warehouse") & ""
	If warehouse_prev <> warehouse or warehouse_prev = "" Then
		response.write "<li class='group'>" & warehouse & " PENDING</li>"
	End If
	warehouse_prev = warehouse

	part = rs("part")
	Description = rs("Description")
	qty = rs("qty")
	id = rs("ID")
	Lft = rs("Lft")
	colour = rs("colour")
	PO = rs("po")

%>
<%
	If warehouse = "SAPA"  Or warehouse = "HYDRO" Or warehouse = "DURAPAINT(WIP)" Then
%>
	<li><a href="stockbyrackedit.asp?ticket=order&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<% Else %>
	<li><a href="stockbyrackedit.asp?ticket=order&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' " %></a></li>
<% End If %>
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
