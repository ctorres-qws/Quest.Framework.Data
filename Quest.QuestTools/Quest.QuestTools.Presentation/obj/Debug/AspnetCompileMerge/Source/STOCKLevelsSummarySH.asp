<!--#include file="dbpath.asp"-->
        <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pendingz)-->
		<!-- Changed August 2015 to add Torbram / Tilton--><!-- Change requested by Shaun Levy, Approved by Jody Cash -->
		<!-- February 2019 - Added USA Option -->
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
        <h1 id="pageTitle">Sheet Stock Level</h1>
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
<li class='group'><a href='StockLevelsSummaryExcel.asp?Job=<%response.write JOB %>' target='_self'>Send to Excel</a></li>         
<form id="Job" title="Stock Level By Job" class="panel" name="job" action="stocklevelsSummarySH.asp" method="GET" target="_self" >
        <h2>Select Job</h2>
  <fieldset>
            <div class="row">
                <label>Job</label>
				<select name="Job" onchange="this.form.submit()" >
				<% ActiveOnly = True %>
				<option value ="">-</option>
                <!--#include file="Jobslist.inc"-->
				<option value ="ALL">ALL</option>
				rsJob.close
				</select>
            </div>
	</fieldset>		
</form>	

<%

if CountryLocation = "USA" then
	Warehouses = "(I.Warehouse = 'JUPITER')"
else
	Warehouses = "(I.Warehouse = 'NASHUA' or I.Warehouse = 'NPREP' AND I.Warehouse = 'GOREWAY' AND I.Warehouse = 'DURAPAINT' AND I.Warehouse = 'DURAPAINT(WIP)' AND I.Warehouse = 'TILTON' AND I.Warehouse = 'MILVAN')"
end if

Set rs = Server.CreateObject("adodb.recordset")
if Job = "" or Job = "ALL" then
	strSQL = "SELECT * FROM Y_INV AS I INNER JOIN Y_MASTER AS M on I.part = M.Part WHERE " & Warehouses & " AND M.InventoryType = 'Sheet' order by M.Part ASC, I.WIDTH ASC, I.HEIGHT ASC"
else

	strSQL = "SELECT * FROM Y_INV AS I INNER JOIN Y_MASTER AS M on I.part = M.Part WHERE " & Warehouses & " AND M.InventoryType = 'Sheet' AND ((Colour = 'Mill' AND Allocation Like '%" & Job &"%' ) OR I.Colour LIKE '%" & Job & "%') order by M.Part ASC , I.WIDTH ASC, I.HEIGHT ASC"
	end if
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	


if job = "" then
job = "ALL"
end if
response.write "<li class='group'>Sheet Stock - Mill/Painted " & Job & " </li>"

if CountryLocation = "USA" then
	response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Size</th><th>Total</th><th>Jupiter Mill</th><th>" & Job & ": Jupiter Painted </th></tr>"
else
	response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Size</th><th>Total</th><th>Nashua Mill</th><th>" & Job & ": Nashua Painted </th><th>Goreway Mill</th><th>" & Job & ": Goreway Painted </th><th>Durapaint Mill</th><th>" & Job & ": Durapaint Painted </th><th>Durapaint(WIP) Mill</th><th>" & Job & ": Durapaint(WIP) Painted </th><th>Tilton Mill</th><th>" & Job & ": Tilton Painted </th><th>Milvan Mill</th><th>" & Job & ": Milvan Painted </th></tr>"
end if

Part = "0"
Size = "0"
		'Goreway and Goreway Painted
	Gqty = 0
	GAqty = 0
		'Nashua and Nashua Painted
	Nqty = 0
	NAqty = 0	
	
		'Durapaint and Durapaint Painted
	Dqty = 0
	DAqty = 0
		'Durapaint(WIP) and Durapaint(WIP) Painted
	DWqty = 0
	DWAqty = 0
		'Tilton Painted
	Tiqty = 0
	TiAqty = 0
	'Milvan Painted
	Mqty = 0
	MAqty = 0
		'Jupiter Painted
	Jqty = 0
	JAqty = 0
	
if not rs.eof then
	rs.movefirst
end if 
	do while not rs.eof
	itemlast = 0
	LastPart = Part
	LastSize = Size
	LastSWidth = SWidth
	LastSHeight = SHeight

	Part= RS("Part") 
	SWidth = RS("Width")
	SHeight = RS("Height")
	Size = RS("Width") & " X " & RS("Height")


	
	
	if LastPart = Part and LastSize = Size then
	' Same Item and Size   (Add to Counter, Do not Display)
		Select Case RS("WAREHOUSE")
		CASE "GOREWAY"
			if rs("colour") = "Mill" then
					Gqty = rs("Qty") + Gqty
				else
					GAqty = rs("Qty") + GAqty
				end if
		CASE "JUPITER"
			if rs("colour") = "Mill" then
					Jqty = rs("Qty") + Jqty
				else
					JAqty = rs("Qty") + JAqty
				end if
		CASE "MILVAN"
			if rs("colour") = "Mill" then
					Mqty = rs("Qty") + Mqty
				else
					MAqty = rs("Qty") + MAqty
				end if
				
		CASE "NASHUA","NPREP"
			if rs("colour") = "Mill" then
					Nqty = rs("Qty") + Nqty
				else
					NAqty = rs("Qty") + NAqty
				end if
		CASE "DURAPAINT"
			if rs("colour") = "Mill" then

				Dqty = rs("Qty") + Dqty
			else
				DAqty = rs("Qty") + DAqty
			end if

		CASE "DURAPAINT(WIP)"
			if rs("colour") = "Mill" then

				DWqty = rs("Qty") + DWqty
			else
				DWAqty = rs("Qty") + DWAqty
			end if

		CASE "TILTON"
		if rs("colour") = "Mill" then

				Tiqty = rs("Qty") + Tiqty
			else
				TiAqty = rs("Qty") + TiAqty
			end if		
		End Select
	itemlast = 1
	else
		'New Item or Size	(Display old, Restart Counter)
		
if  CountryLocation = "USA" then

		if Jqty + JAqty = 0 then
			else
	
				response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=sheet&part=" & LastPart & "&Swidth=" & LastSWidth & "&SHeight=" & LastSHeight & "' target='_self'>" & LastPart & "</a></td>"
				response.write "<td> " & LastSize  & "</td>"
				response.write "<td><b> " & Jqty+ JAqty & "</b></td>"
				response.write "<td>" & Jqty & "</td><td>" & JAqty & "</td>"
				response.write "</tr>"
		end if
else

		if Gqty+ GAqty + Nqty + NAqty + Dqty+ DAqty + DWqty + DWAqty + Tiqty + TiAqty = 0 then
			else
	
				response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=sheet&part=" & LastPart & "&Swidth=" & LastSWidth & "&SHeight=" & LastSHeight & "' target='_self'>" & LastPart & "</a></td>"
				response.write "<td> " & LastSize  & "</td>"
				response.write "<td><b> " & Gqty+ GAqty + Nqty + NAqty + Dqty+ DAqty + DWqty + DWAqty + Tiqty + TiAqty + Mqty + MAqty & "</b></td>"
				response.write "<td>" & Nqty & "</td><td>" & NAqty & "</td>"
				response.write "<td>" & Gqty & "</td><td>" & GAqty & "</td>"
				response.write "<td>" & Dqty & "</td><td>" & DAqty & "</td>"
				response.write "<td>" & DWqty & "</td><td>" & DWAqty & "</td>"
				response.write "<td>" & Tiqty & "</td><td>" & TiAqty & "</td>"
				response.write "<td>" & Mqty & "</td><td>" & MAqty & "</td>"
				response.write "</tr>"
				
		end if
end if	
	
	
	itemlast = 1
	Gqty = 0
	GAqty = 0
	Jqty = 0
	JAqty = 0
	Nqty = 0
	NAqty = 0
	Dqty = 0
	DAqty = 0
	DWqty = 0
	DWAqty = 0
	Tiqty = 0
	TiAqty = 0
	Mqty = 0
	MAqty = 0	
	Select Case RS("WAREHOUSE")
		CASE "GOREWAY"
			if rs("colour") = "Mill" then
					Gqty = rs("Qty") + Gqty
				else
					GAqty = rs("Qty") + GAqty
				end if
		CASE "JUPITER"
			if rs("colour") = "Mill" then
					Jqty = rs("Qty") + Jqty
				else
					JAqty = rs("Qty") + JAqty
				end if		
		CASE "NASHUA","NPREP"
			if rs("colour") = "Mill" then
					Nqty = rs("Qty") + Nqty
				else
					NAqty = rs("Qty") + NAqty
				end if
		CASE "DURAPAINT"
			if rs("colour") = "Mill" then

				Dqty = rs("Qty") + Dqty
			else
				DAqty = rs("Qty") + DAqty
			end if
		CASE "MILVAN"
			if rs("colour") = "Mill" then

				Mqty = rs("Qty") + Mqty
			else
				MAqty = rs("Qty") + MAqty
			end if
			
		CASE "DURAPAINT(WIP)"
			if rs("colour") = "Mill" then

				DWqty = rs("Qty") + DWqty
			else
				DWAqty = rs("Qty") + DWAqty
			end if

		CASE "TILTON"
		if rs("colour") = "Mill" then

				Tiqty = rs("Qty") + Tiqty
			else
				TiAqty = rs("Qty") + TiAqty
			end if		
		End Select
	
	end if 

	rs.movenext
	loop
	
	if itemlast = 1 then
	
		if  CountryLocation = "USA" then

				if Jqty + JAqty = 0 then
					else
			
					response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=sheet&part=" & LastPart & "&Swidth=" & LastSWidth & "&SHeight=" & LastSHeight & "' target='_self'>" & LastPart & "</a></td>"
					response.write "<td> " & LastSize  & "</td>"
					response.write "<td><b> " & Jqty+ JAqty & "</b></td>"
					response.write "<td>" & Jqty & "</td><td>" & JAqty & "</td>"
					response.write "</tr>"
					itemlast = 1
				end if
		else
	
				if Gqty+ GAqty + Nqty + NAqty + Dqty+ DAqty + DWqty + DWAqty + Tiqty + TiAqty + Mqty + MAqty = 0 then
				else
				
					response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=sheet&part=" & Part & "&Swidth=" & SWidth & "&SHeight=" & SHeight & "' target='_self'>" & Part & "</a></td>"
					response.write "<td> " & Size  & "</td>"
					response.write "<td><b> " & Gqty+ GAqty + Nqty + NAqty + Dqty+ DAqty + DWqty + DWAqty + Tiqty + TiAqty & "</b></td>"
					response.write "<td>" & Nqty & "</td><td>" & NAqty & "</td>"
					response.write "<td>" & Gqty & "</td><td>" & GAqty & "</td>"
					response.write "<td>" & Dqty & "</td><td>" & DAqty & "</td>"
					response.write "<td>" & DWqty & "</td><td>" & DWAqty & "</td>"
					response.write "<td>" & Tiqty & "</td><td>" & TiAqty & "</td>"
					response.write "<td>" & Mqty & "</td><td>" & MAqty & "</td>"

					response.write "</tr>"
					itemlast = 1
				end if
		end if
	end if
	
rs.close
set rs = nothing


DBConnection.close
set DBConnection=nothing
response.write "</table></li>"

%>

   
            
   </ul>
</body>
</html>
