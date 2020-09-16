<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
			   <!-- Changed August 2015 to  add Torbram / Tilton-->
				<!-- Change requested by Shaun Levy, Approved by Jody Cash -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>
    
      
  
<%

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=StockSummary" & Date() & ".xls"
Job = Request.QueryString("Job")
%>  

<ul id="screen1" title="Stock Level <% response.write ": " & Job %>" selected="true">            
<%



'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill/Painted " & Job & " </li>"
if job = "" then
job = "ALL"
end if
response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Total</th><th>Min level</th><th>Goreway Mill</th><th>" & Job & ": Goreway Allocated </th><th>Durapaint Mill</th><th>" & Job & ": Durapaint Allocated </th><th>" & Job & ": Durapaint(WIP) Allocated </th><th>Horner Mill</th><th>" & Job & ": Horner Allocated </th><th>HYDRO Mill</th><th>" & Job & ": HYDRO Allocated </th><th>Painted: " & Job & "</th><th>Nashua Mill</th><th>" & Job & ": Nashua Allocated </th><th>Tilton Mill</th><th>" & Job & ": Tilton Allocated </th><th>Milvan Mill</th><th>" & Job & ": Milvan Allocated </th><th>Painted: " & Job & "</th><th>Pending</th></tr>"
'<th>Durapaint(WIP) Mill</th>
rs2.movefirst
	do while not rs2.eof
	'Goreway and Goreway Allocated
	Gqty = 0
	GAqty = 0
	'Durapaint and Durapaint Allocated
	Dqty = 0
	DAqty = 0
	'Durapaint(WIP) and Durapaint(WIP) Allocated
	' Removed this one, it should be always be zero DWqty = 0
	DWAqty = 0
	'Sapa and Sapa Allocated
	Sqty = 0
	SAqty = 0
	'Horner and Horner Allocated
	Hqty = 0
	HAqty = 0
	'Nashua and Nashua Allocated
	Nqty = 0
	NAqty = 0
	'Tilton and Tilton Allocated
	Tiqty = 0
	TiAqty = 0
	'Milvan and Milvan Allocated
	Mqty = 0
	MAqty = 0
	
	partqty2 = 0
	partqty3 = 0
	
Set rs = Server.CreateObject("adodb.recordset")
if Job = "" or Job = "ALL" then
	strSQL = "SELECT * FROM Y_INV WHERE (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE ((Colour = 'Mill' AND Allocation Like '%" & Job &"%' ) OR Colour LIKE '%" & Job & "%')  And (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"
end if
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection	
	if rs.eof then
	else
	rs.movefirst
	do while not rs.eof
	Select Case RS("WAREHOUSE")
	CASE "GOREWAY"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Gqty = rs("Qty") + Gqty
			else
				GAqty = rs("Qty") + GAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "DURAPAINT"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Dqty = rs("Qty") + Dqty
			else
				DAqty = rs("Qty") + DAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "DURAPAINT(WIP)"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				'DWqty = rs("Qty") + DWqty
			else
				DWAqty = rs("Qty") + DWAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "HORNER"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Hqty = rs("Qty") + Hqty
			else
				HAqty = rs("Qty") + HAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "SAPA","HYDRO"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Sqty = rs("Qty") + Sqty
			else
				SAqty = rs("Qty") + SAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "NASHUA","NPREP"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Nqty = rs("Qty") + Nqty
			else
				NAqty = rs("Qty") + NAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "TILTON"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Tiqty = rs("Qty") + Tiqty
			else
				TiAqty = rs("Qty") + TiAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if
	CASE "MILVAN"
		if rs("colour") = "Mill" then
			if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
				Mqty = rs("Qty") + Mqty
			else
				MAqty = rs("Qty") + MAqty
			end if
		else
					partqty2 = rs("Qty") + partqty2
		end if	
	
	CASE "JUPITER", "JUPITER PRODUCTION"	
	CASE Else
		partqty3 = rs("Qty") + partqty3

	End Select

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if Gqty + Dqty + DWqty + Hqty + Sqty + Nqty + Tiqty + Mqty + partqty3 < rs2("MinLevel") AND partqty2 + partqty3 < rs2("MinLevel") then
		MinLevelAlert = "Below"
	end if


		if partqty2 = 0 and Gqty = 0 and GAqty = 0 and Dqty = 0 and DAqty = 0 and DWqty = 0 and DWAqty = 0 and Hqty = 0 and HAqty = 0 and Sqty = 0 and SAqty = 0 and Nqty = 0 and NAqty = 0 and Tiqty = 0 and TiAqty = 0 and Mqty = 0 and MAqty = 0 then
		else
			response.write "<tr><td>" & rs2("part") & "</td>"
			response.write "<td>" & rs2("description") & "</td>"
			response.write "<td> " & Gqty+ GAqty + Dqty + DAqty + DWAqty + Hqty + HAqty + Sqty + SAqty + Nqty + NAqty + Tiqty + TiAqty + Mqty + MAqty +  partqty2 & "</td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & rs2("MinLevel") & "</font></td>"
			else
				response.write "<td>" & rs2("MinLevel") & "</td>"
		
			end if
			response.write "<td>" & Gqty & "</td><td>" & GAqty & "</td>"
			response.write "<td>" & Dqty & "</td><td>" & DAqty & "</td>"
			response.write "<td>" & DWAqty & "</td>"
			response.write "<td>" & Hqty & "</td><td>" & HAqty & "</td>"
			response.write "<td>" & Sqty & "</td><td>" & SAqty & "</td>"
			response.write "<td>" & Nqty & "</td><td>" & NAqty & "</td>"
			response.write "<td>" & Tiqty & "</td><td>" & TiAqty & "</td>"
			response.write "<td>" & Mqty & "</td><td>" & MAqty & "</td>"
			response.write "<td> " & partqty2 & "</td>"
			response.write "<td> " & partqty3 & "</td>"
			
			response.write "</tr>"
			'<td>" & DWqty & "</td>
		end if 

	end if
	
rs.close
set rs = nothing
'end if
rs2.movenext
loop
response.write "</table></li>"

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>          
 </ul>
</body>
</html>


