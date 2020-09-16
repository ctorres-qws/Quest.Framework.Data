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

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

        <ul id="Profiles" title="Pending Stock" selected="true">
         <li class="group"><a href="stockpendingtable.asp?part=<%response.write part%>" target="_self" >Stock Pending (Row Form) - Switch to Table Form</a></li>
<%
	'New Code to Write filter by all Warehouses in one location instead of individually.
	'Old code set to if True = False
	Dim a_Warehouses
	a_Warehouses = Array("SAPA","HYDRO","DURAPAINT(WIP)","DEPENDABLE","EXTAL SEA","KEYMARK","TILTON(WIP)","CAN-ART","APEL","METRA","EXTRUDEX")
	For i = 0 to UBound(a_Warehouses)
		Response.Write "<li class='group'>" & a_Warehouses(i) & " PENDING</li>"
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT yM.Description, yI.* FROM Y_Inv yI LEFT JOIN Y_Master yM ON yM.Part = yI.Part WHERE WAREHOUSE ='" + a_Warehouses(i) & "' ORDER BY WAREHOUSE, yI.PART"
		rs.Cursortype = GetDBCursorType
		rs.Locktype = GetDBLockType
		rs.Open strSQL, DBConnection

		Do While Not rs.eof
			Description = rs("Description") & ""
			If Description = "" Then
				Description = "N/A"
			End If

			part = rs("Part")
			qty = rs("qty")
			id = rs("ID")
			Lft = rs("Lft")
			colour = rs("colour")
			PO = rs("po")
			Datein = rs("Datein")
			Allocation = " - Allocated to: " & rs("Allocation")
			Select Case (a_Warehouses(i))
				Case "EXTAL SEA", "KEYMARK", "CAN-ART", "METRA", "EXTRUDEX"
					Allocation = ""
			End Select
%>
			<li><a href="stockbyrackedit.asp?ticket=order&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' " & Allocation & " Order Date: " & Datein %></a></li>
<%
			rs.MoveNext
		Loop
		rs.Close

	Next

set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>

</body>
</html>
