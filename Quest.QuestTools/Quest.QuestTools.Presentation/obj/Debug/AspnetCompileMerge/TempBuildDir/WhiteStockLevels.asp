<!--#include file="dbpath.asp"-->
              <!-- New Report to show White Stock levels (FOr Mary Darnell, Written by Michael Bernholtz, November 2016-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>White Stock Levels</title>
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

    <div class="toolbar" >
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>

<ul id="screen1" title="Stock Level"  selected="true">           
<%

	response.write "<li class='group'>White Stock by Die  (Goreway, Durapaint, Horner, Nashua, Torbram, HYDRO)</li>"

	response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Stock Length</th><th>In Warehouse</th><th>HYDRO Pending</th></tr>"

	Set rs = Server.CreateObject("adodb.recordset")
	'strSQL = "SELECT * FROM Y_INV WHERE (Warehouse ='GOREWAY' or Warehouse ='SAPA' or Warehouse ='TORBRAM' or Warehouse ='NASHUA' or Warehouse ='DURAPAINT' or Warehouse ='DURAPAINT(WIP)' or Warehouse ='HORNER') And Part = '" & rs2("Part") & "' AND Colour = 'White' order by LFT ASC"
	'strSQL = "SELECT * FROM Y_Master yM LEFT JOIN Y_INV yI ON (yM.Part = yI.Part AND yI.Warehouse IN ('GOREWAY','SAPA','TORBRAM','NASHUA','DURAPAINT','DURAPAINT(WIP)','HORNER') AND yI.Colour = 'White') WHERE INVENTORYTYPE = 'Extrusion' ORDER BY LFT ASC"
	'strSQL = "SELECT yM.Part, yM.Description, yI.IQty as Qty, yI.Lft, yI.Warehouse as Warehouse FROM Y_Master yM LEFT JOIN ( SELECT Sum(Qty) as IQty, Warehouse, Part, Lft FROM y_Inv WHERE Warehouse IN ('GOREWAY','SAPA','TORBRAM','NASHUA','DURAPAINT','DURAPAINT(WIP)','HORNER') AND Colour = 'White' GROUP By Part, Warehouse, Lft ) yI ON yI.Part = yM.Part WHERE INVENTORYTYPE = 'Extrusion' ORDER BY yM.Part, LFT ASC"
	strSQL = "SELECT yM.Part, yM.Description, yI.IQty as Qty, yI.Lft, yI.Warehouse as Warehouse FROM Y_Master yM LEFT JOIN ( SELECT Sum(Qty) as IQty, Warehouse, Part, Lft FROM y_Inv INNER JOIN y_Color yC ON y_Inv.COLOUR = yC.PROJECT WHERE Warehouse IN ('GOREWAY','SAPA','HYDRO','TORBRAM','NASHUA','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') AND yC.CODE = 'K1285' GROUP By Part, Warehouse, Lft ) yI ON yI.Part = yM.Part WHERE INVENTORYTYPE = 'Extrusion' ORDER BY yM.Part, LFT ASC"

	rs.Cursortype = GetDBCursorType
	rs.Locktype = GetDBLockType
	rs.Open strSQL, DBConnection
	QTY = 0
	SQTY = 0

	str_Length = 0.0
	str_Part = ""

	On Error Resume Next

	Do While Not rs.eof

		If str_Length = FixFloat(rs("LFT") & "") And str_Part = rs("Part") & "" Then
			If RS("Warehouse") = "SAPA" or RS("Warehouse") = "HYDRO" Then
				SQTY = SQTY + rs("QTY")
			Else
				Qty = QTY + rs("QTY")
			End If
		Else
			If QTY>0 Or SQTY > 0 Then
				response.write "<tr>"
				response.write "<td>" & str_Part & "</td>"
				response.write "<td>" & str_Description & "</td>"
				response.write "<td>" & str_Length & "</td>"
				response.write "<td>" & QTY & "</td>"
				response.write "<td>" & SQTY & "</td>"
				response.write "</tr>"
				SQTY = 0
				QTY = 0
			End If

			If RS("Warehouse") = "SAPA" or RS("Warehouse") = "HYDRO" Then
				SQTY = rs("QTY")
				QTY = 0
			Else
				QTY = rs("QTY")
				SQTY= 0
			End If

		End If

		str_Part = rs("Part")
		str_Description = rs("Description")
		str_Length = rs("Lft")

		rs.MoveNext
	Loop

	If QTY > 0 Then
		response.write "<tr>"
		response.write "<td>" & str_Part & "</td>"
		response.write "<td>" & str_Description & "</td>"
		response.write "<td>" & str_Length & "</td>"
		response.write "<td>" & QTY & "</td>"
		response.write "<td>" & SQTY & "</td>"
		response.write "</tr>"
		response.write "</tr>"
	End If

rs.close
set rs = nothing
%>

   </ul>
</body>
</html>

<%

DBConnection.close
set DBConnection=nothing
%>

