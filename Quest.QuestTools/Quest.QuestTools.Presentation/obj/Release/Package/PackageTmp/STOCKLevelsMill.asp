<!--#include file="dbpath.asp"-->
              <!-- New Report to show Mill Stock levels (ignoring all painted - Requested by Shaun Levy, Written by Michael Bernholtz, December 2015-->
			  <!-- Updated October 2018 - Tilton and Metra added as column-->
			  <!-- February 2019 - Added USA Option -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Mill Stock Levels</title>
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

 <div class="toolbar">
        <h1 id="pageTitle">Mill Stock Levels</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>
<ul id="screen1" title="Mill Extrusion"  selected="true">           


<%



'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER where INVENTORYTYPE = 'Extrusion' order by PART ASC"
'Get a Record Set
'    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill</li>"
response.write "<li class='group'>Other Pending includes: Dependable, Extal Sea, Keymark, TechnoForm </li>"
if job = "" then
job = "ALL"
end if

if CountryLocation = "USA" then

response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Available</th><th>Min level</th><th>Jupiter Mill</th></tr>"
else
response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Available</th><th>Min level</th><th>Goreway Mill</th><th>Durapaint Mill</th><th>Horner Mill</th>"
response.write "<th>Nashua Mill</th><th>Can-Art Mill</th><th>HYDRO Mill</th><th>Durapaint(WIP) Mill</th><th>APEL Mill</th><th>Metra Mill</th><th>Tilton Mill</th><th>Milvan Mill</th><th>Other Pending</th></tr>"

end if
'rs2.movefirst
'	do while not rs2.eof
	'MillValues for Each Warehouse
	Gqty = 0
	Jqty = 0
	Dqty = 0
	Nqty = 0
	Mqty = 0
	Miqty = 0
	DWqty = 0
	Sqty = 0
	Hqty = 0
	Toqty = 0
	Tiqty = 0
	Caqty = 0
	Aqty = 0
	partqty3 = 0

	Set rs = Server.CreateObject("adodb.recordset")
	if Job = "" or Job = "ALL" then
		'strSQL = "SELECT * FROM Y_INV WHERE (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' AND Colour = 'Mill' order by Part ASC"
		strSQL = "SELECT yM.Part, yM.MinLevel, yM.Description, yI.Qty, yI.Warehouse FROM Y_MASTER yM LEFT JOIN y_Inv yI ON (yI.Part = yM.Part AND Colour = 'Mill' AND yI.Warehouse <> 'WINDOW PRODUCTION' AND yI.Warehouse <> 'COM PRODUCTION' AND yI.Warehouse <> 'SCRAP' AND yI.Warehouse <> 'JUPITER PRODUCTION') WHERE INVENTORYTYPE = 'Extrusion' order by yM.PART ASC "
	else
		'strSQL = "SELECT * FROM Y_INV WHERE Colour = 'Mill'  And (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Part ASC"
		strSQL = "SELECT yI.*, yM.MinLevel, yM.Description FROM Y_MASTER yM LEFT JOIN y_Inv yI ON (yI.Part = yM.Part AND Colour = 'Mill' And (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP' AND Warehouse <> 'JUPITER PRODUCTION')) WHERE INVENTORYTYPE = 'Extrusion' ORDER BY Part ASC"
	end if

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	

	do while not rs.eof


	If str_Part <> rs("Part") Then

		DisplayRow

		'MillValues for Each Warehouse
		Gqty = 0
		Jqty = 0
		Dqty = 0
		Nqty = 0
		Mqty = 0
		Miqty = 0
		DWqty = 0
		Sqty = 0
		Hqty = 0
		Toqty = 0
		Tiqty = 0
		Caqty = 0
		Aqty = 0
		partqty3 = 0

	End If

	str_Part = rs("Part")
	str_Desc = rs("Description")
	str_MinLevel = rs("MinLevel")

	Select Case UCASE(RS("WAREHOUSE"))
	CASE "GOREWAY"
				Gqty = rs("Qty") + Gqty
	CASE "JUPITER"
				Jqty = rs("Qty") + Jqty
	CASE "DURAPAINT"
				Dqty = rs("Qty") + Dqty
	CASE "DURAPAINT(WIP)"
				DWqty = rs("Qty") + DWqty
	CASE "HORNER"
				Hqty = rs("Qty") + Hqty
	CASE "NASHUA","NPREP"
				Nqty = rs("Qty") + Nqty
	CASE "METRA"
				Mqty = rs("Qty") + Mqty
	CASE "MILVAN"
				Miqty = rs("Qty") + Miqty
	CASE "SAPA","HYDRO"
				Sqty = rs("Qty") + Sqty
	CASE "CAN-ART"
				Caqty = rs("Qty") + Caqty 
	CASE "TILTON", "TILTON(WIP)"
				Tiqty = rs("Qty") + Tiqty
	CASE "APEL"
				Aqty = rs("Qty") + Aqty				
	CASE Else
		partqty3 = rs("Qty") + partqty3

	End Select

	rs.movenext
	loop

	DisplayRow
	
rs.close
set rs = nothing



%>
</table></li>

   </ul>
</body>
</html>

<%

DBConnection.close
set DBConnection=nothing
%>
<%
	Function DisplayRow
if CountryLocation = "USA" then
		MinLevelAlert = ""
		if Jqty  < str_MinLevel then
			MinLevelAlert = "Below"
		end if

		if Jqty = 0 and str_MinLevel = 0 then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & str_Part & "' target='_self'>" & str_Part & "</a></td>"
			response.write "<td>" & str_Desc & "</td>"
			response.write "<td><B> " &  Jqty & "</B></td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & str_MinLevel & "</font></td>"
			else
				response.write "<td>" & str_MinLevel & "</td>"
			end if
			response.write "<td title='Jupiter Mill'><B>" & Jqty & "</B></td>"
			response.write "</tr>"

		end if 
else
		MinLevelAlert = ""
		if Gqty + Nqty + Dqty + Hqty + Caqty  < str_MinLevel then
			MinLevelAlert = "Below"
		end if

		if Gqty = 0 and Dqty = 0 and Nqty = 0 and DWqty = 0 and Hqty = 0  and Sqty = 0  and Caqty = 0 and Tiqty = 0 and Miqty = 0 and Aqty = 0 and Mqty = 0 and str_MinLevel = 0 then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & str_Part & "' target='_self'>" & str_Part & "</a></td>"
			response.write "<td>" & str_Desc & "</td>"
			response.write "<td><B> " &  Gqty + Dqty + Hqty + Nqty + Miqty & "</B></td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & str_MinLevel & "</font></td>"
			else
				response.write "<td>" & str_MinLevel & "</td>"
			end if
			response.write "<td title='Goreway Mill'><B>" & Gqty & "</B></td>"
			response.write "<td title='Durapaint Mill'><B>" & Dqty & "</B></td>"
			response.write "<td title='Horner'><B>" & Hqty & "</B></td>"
			response.write "<td title='Nashua'><B>" & Nqty & "</B></td>"
			response.write "<td title='Can-Art'>" & Caqty & "</td>"
			response.write "<td title='HYDRO'>" & Sqty & "</td>"
			response.write "<td title='Durapaint WIP'>" & DWqty & "</td>"
			response.write "<td title='APEL'>" & Aqty & "</td>"
			response.write "<td title='Metra'>" & Mqty & "</td>"
			response.write "<td title='Tilton'>" & Tiqty & "</td>"
			response.write "<td title='Tilton'>" & Miqty & "</td>"
			response.write "<td title='Other'> " & partqty3 & "</td>"
			response.write "</tr>"

		end if 

end if
	End Function

%>
