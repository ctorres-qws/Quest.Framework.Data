<!--#include file="dbpath.asp"-->
    <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
	<!-- Changed August 2015 to  add Torbram / Tilton-->
	<!-- Originally clone of Extrusion, Reformatted to be Plastic specific December 2016, Mary Darnell-->
	<!-- February 2019 - Added USA Option -->
	<!-- October 2019 - Corrected USA error - not reseting JQty Counter - this has been wrong since Feb, but no one noticed - Michael Bernholtz-->
	
				
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
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
        <h1 id="pageTitle">Plastic Stock Level</h1>
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
<%
Job = Request.QueryString("Job")
%>
<ul id="screen1" title="Stock Level <% response.write ": " & Job %>" selected="true">        

<%

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER where INVENTORYTYPE = 'Plastic' order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Plastic Stock </li>"

if CountryLocation = "USA" then
	response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Total</th><th>Min level</th><th>Length</th><th>Jupiter</th><th>Pending</th></tr>"
else	
	response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Total</th><th>Min level</th><th>Length</th><th>Goreway</th><th>Horner</th><th>Nashua</th><th>Milvan</th><th>Dura/Tilton/Other</th><th>Technoform</th><th>Metra</th><th>Pending</th></tr>"
'<th>Durapaint(WIP) Mill</th>
end if

rs2.movefirst

	do while not rs2.eof 
	i = 1
do until i>9
SELECT CASE  i
case 1
	i = 2
	PLLength = 16
	PLMin = 16
case 2
	i = 3
	PLLength = 17
	PLMin = 17
case 3
	i = 4
	PLLength = 18
	PLMin = 18
case 4
	i = 5
	PLLength = 19
	PLMin = 19
case 5
	i = 6
	PLLength = 20
	PLMin = 20
case 6
	i = 7
	PLLength = 21
	PLMin = 21
case 7
	i = 8
	PLLength = 21.5
	PLMin = 22
case 8
	PLLength = 22
	PLMin = 22
	i = 10
end select
	
	'Goreway
	Gqty = 0
	'Horner 
	Hqty = 0
	'Nashua
	Nqty = 0
	'Technoform 
	Teqty = 0
	'Metra
	Mqty = 0
	'Milvan
	Miqty = 0
	'Jupiter
	Jqty = 0
		
	partqty2 = 0
	partqty3 = 0
	
If CountryLocation = "USA" then
	WareHouses = "(Warehouse = 'JUPITER')"
else
	WareHouses = "(Warehouse = 'GOREWAY' OR Warehouse = 'DURAPAINT' OR Warehouse = 'NASHUA' OR Warehouse = 'NPREP' OR Warehouse = 'Horner' OR Warehouse = 'DURAPAINT(WIP)' OR Warehouse = 'CAN-ART' OR Warehouse = 'EXTAL SEA' OR Warehouse = 'MILVAN' OR Warehouse = 'DEPENDABLE' OR Warehouse = 'SAPA' OR Warehouse = 'HYDRO')"
end if
	
Set rs = Server.CreateObject("adodb.recordset")


	strSQL = "SELECT * FROM Y_INV WHERE " & Warehouses & " And Part = '" & rs2("Part") & "' And Lft = " & PLLength & " order by LFT,  Colour ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	

	if rs.eof then
	else
	rs.movefirst
	do while not rs.eof
		
		
		Select Case RS("WAREHOUSE")
		CASE "GOREWAY"
					Gqty = rs("Qty") + Gqty
		CASE "HORNER"
					Hqty = rs("Qty") + Hqty
		CASE "NASHUA","NPREP"
					Nqty = rs("Qty") + Nqty
		CASE "TECHNOFORM"
					Teqty = rs("Qty") + Teqty					
		CASE "METRA"
					Mqty = rs("Qty") + Mqty
		CASE "MILVAN"
					Miqty = rs("Qty") + Miqty	
		CASE "JUPITER"
					Jqty = rs("Qty") + Jqty		
		CASE "TILTON","DURAPAINT","DURAPAINT(WIP)"
			partqty2 = rs("Qty") + partqty2			
		CASE Else
			partqty3 = rs("Qty") + partqty3

		End Select

	rs.movenext
	loop

	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz

	if CountryLocation = "USA" then
	
		MinLevelAlert = ""
		if Jqty + partqty2 < rs2("Min-"& PLMin) AND partqty2 + partqty3 < rs2("Min-"& PLMIN) then
			MinLevelAlert = "Below"
		end if

			if partqty2 = 0 and Jqty = 0 then
			else
				response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=plastic&part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
				response.write "<td>" & rs2("description") & "</td>"
				response.write "<td> " & Jqty + partqty2 & "</td>"
				if MinLevelAlert ="Below" then
					response.write "<td><font color='red'> " & rs2("Min-"& PLMin) & "</font></td>"
				else
					response.write "<td>" & rs2("Min-"& PLMin) & "</td>"
				end if
				response.write "<td>" & PLLength & "</td>"
				response.write "<td>" & Jqty & "</td>"
				response.write "<td>" & partqty2 & "</td>"
				response.write "<td> " & partqty3 & "</td>"
				response.write "</tr>"
			end if 
	else
		
			MinLevelAlert = ""
			if Gqty + Hqty + Nqty + Miqty + partqty2 < rs2("Min-"& PLMin) AND partqty2 + partqty3 < rs2("Min-"& PLMIN) then
				MinLevelAlert = "Below"
			end if

			if partqty2 = 0 and Gqty = 0 and Hqty = 0 and Nqty = 0 and Miqty = 0 and Teqty = 0 and Mqty = 0 then
			else
				response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=plastic&part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
				response.write "<td>" & rs2("description") & "</td>"
				response.write "<td> " & Gqty + Hqty + Nqty  + Miqty + partqty2 & "</td>"
				if MinLevelAlert ="Below" then
					response.write "<td><font color='red'> " & rs2("Min-"& PLMin) & "</font></td>"
				else
					response.write "<td>" & rs2("Min-"& PLMin) & "</td>"
				end if
				response.write "<td>" & PLLength & "</td>"
				response.write "<td>" & Gqty & "</td>"
				response.write "<td>" & Hqty & "</td>"
				response.write "<td>" & Nqty & "</td>"
				response.write "<td>" & Miqty & "</td>"
				response.write "<td>" & partqty2 & "</td>"
				response.write "<td>" & Teqty & "</td>"
				response.write "<td>" & Mqty & "</td>"

				response.write "<td> " & partqty3 & "</td>"

				response.write "</tr>"

			end if 
		end if

rs.close
set rs = nothing
end if
loop
rs2.movenext
loop

response.write "</table></li>"

%>

   </ul>
</body>
</html>

<%

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

