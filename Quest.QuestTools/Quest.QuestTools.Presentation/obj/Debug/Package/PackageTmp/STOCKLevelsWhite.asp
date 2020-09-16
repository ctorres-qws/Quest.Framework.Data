<!--#include file="dbpath.asp"-->
              <!-- Based on the Stock Levels Mill report, Exact Same report for White-->
			  
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



'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER where INVENTORYTYPE = 'Extrusion' order by PART ASC"
'Get a Record Set
'    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - White </li>"
response.write "<li class='group'>Other Pending includes: Dependable, Extal Sea, Keymark, Metra, Tilton, TechnoForm </li>"

response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Available</th><th>Goreway White</th><th>Durapaint White</th><th>Horner White</th><th>Nashua White</th><th>Milvan White</th><th>Can-Art White</th><th>HYDRO White</th><th>Durapaint(WIP) White</th><th>APEL White</th><th>Other Pending</th></tr>"
	'WhiteValues for Each Warehouse
	Gqty = 0
	Dqty = 0
	Nqty = 0
	DWqty = 0
	Sqty = 0
	Hqty = 0
	Toqty = 0
	Caqty = 0
	Aqty = 0
	Mqty = 0
	partqty3 = 0

	Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT yM.Part, yM.MinLevel, yM.Description, yI.Qty, yI.Warehouse FROM Y_MASTER yM LEFT JOIN y_Inv yI ON (yI.Part = yM.Part AND Colour = 'White' AND yI.Warehouse <> 'WINDOW PRODUCTION' AND yI.Warehouse <> 'COM PRODUCTION' AND yI.Warehouse <> 'SCRAP' AND yI.Warehouse <> 'JUPITER' AND yI.Warehouse <> 'JUPITER PRODUCTION') WHERE INVENTORYTYPE = 'Extrusion' order by yM.PART ASC "
rs.Cursortype = 2
rs.Locktype = 3

rs.Open strSQL, DBConnection	

	do while not rs.eof


	If str_Part <> rs("Part") Then

		DisplayRow

		'WhiteValues for Each Warehouse
		Gqty = 0
		Dqty = 0
		Nqty = 0
		Mqty = 0
		DWqty = 0
		Sqty = 0
		Hqty = 0
		Toqty = 0
		Caqty = 0
		Aqty = 0
		partqty3 = 0

	End If

	str_Part = rs("Part")
	str_Desc = rs("Description")
	

	Select Case RS("WAREHOUSE")
	CASE "GOREWAY"
		Gqty = rs("Qty") + Gqty
	CASE "DURAPAINT"
		Dqty = rs("Qty") + Dqty
	CASE "DURAPAINT(WIP)"
		DWqty = rs("Qty") + DWqty
	CASE "HORNER"
		Hqty = rs("Qty") + Hqty
	CASE "NASHUA","NPREP"
		Nqty = rs("Qty") + Nqty
	CASE "SAPA", "HYDRO"
		Sqty = rs("Qty") + Sqty
	CASE "CAN-ART"
		Caqty = rs("Qty") + Caqty
	CASE "APEL"
		Aqty = rs("Qty") + Aqty
	CASE "MILVAN"
		Mqty = rs("Qty") + Mqty
	CASE "JUPITER","JUPITER PRODUCTION"
	CASE Else
		partqty3 = rs("Qty") + partqty3
	End Select

	rs.movenext
	loop

	DisplayRow

	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz

rs.close
set rs = nothing

'rs2.movenext
'loop

%>
<tr><th>Stock</th><th>Description</th><th>Available</th><th>Min level</th><th>Goreway White</th><th>Durapaint White</th><th>Horner White</th><th>Nashua  White</th><th>Milvan White</th><th>Can-Art White</th><th>HYDRO White</th><th>Durapaint(WIP) White</th><th>APEL White</th><th>Other Pending</th></tr>
</table></li>

   </ul>
</body>
</html>

<%

'rs2.close
'set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>
<%
	Function DisplayRow


		if Gqty = 0 and Dqty = 0 and Nqty = 0 and DWqty = 0 and Hqty = 0  and Sqty = 0  and Caqty = 0 and Aqty = 0 and Mqty = 0 then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & str_Part & "' target='_self'>" & str_Part & "</a></td>"
			response.write "<td>" & str_Desc & "</td>"
			response.write "<td><B> " &  Gqty + Dqty + Hqty + Nqty + Mqty & "</B></td>"
			response.write "<td title='Goreway White'><B>" & Gqty & "</B></td>"
			response.write "<td title='Durapaint White'><B>" & Dqty & "</B></td>"
			response.write "<td title='Horner'><B>" & Hqty & "</B></td>"
			response.write "<td title='Nashua'><B>" & Nqty & "</B></td>"
			response.write "<td title='Nashua'><B>" & Mqty & "</B></td>"
			response.write "<td title='Can-Art'>" & Caqty & "</td>"
			response.write "<td title='HYDRO'>" & Sqty & "</td>"
			response.write "<td title='Durapaint WIP'>" & DWqty & "</td>"
			response.write "<td title='APEL'>" & Aqty & "</td>"
			response.write "<td title='Other'> " & partqty3 & "</td>"
			response.write "</tr>"

		end if 

	End Function

%>
