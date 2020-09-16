<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
				<!-- Change requested by Shaun Levy, Approved by Jody Cash -->
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
        <h1 id="pageTitle">Stock Levels</h1>
        <% 
		
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write HomeSite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>
    
      
  
<%
Job = Request.QueryString("Job")
%>  
<ul id="screen1" title="Stock Level <% response.write ": " & Job %>" selected="true">            
<form id="Job" title="Stock Level By Job" class="panel" name="job" action="stocklevels.asp" method="GET" target="_self" >
        <h2>Select Job</h2>
  <fieldset>
            <div class="row">
                <label>Job</label>
				<select name="Job" onchange="this.form.submit()" >
				<% ActiveOnly = True %>
				<option value ="">-</option
                <!--#include file="Jobslist.inc"-->
				rsJob.close
				</select>
            </div>
	</fieldset>		
</form>	

<%
If CountryLocation = "USA" then
	WareHouses = "(Warehouse = 'JUPITER')"
else
	WareHouses = "(Warehouse = 'GOREWAY' OR Warehouse = 'DURAPAINT' OR Warehouse = 'NASHUA' OR Warehouse = 'NPREP' OR Warehouse = 'Horner' OR Warehouse = 'DURAPAINT(WIP)' OR Warehouse = 'CAN-ART' OR Warehouse = 'EXTAL SEA' OR Warehouse = 'DEPENDABLE' OR Warehouse = 'SAPA' OR Warehouse = 'HYDRO')"
end if


if Job = "" or Job = "ALL" then
	strSQL = "SELECT * FROM Y_INV WHERE " & WareHouses & " order by Colour ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE (((Colour = 'Mill' AND Allocation Like '%" & Job &"%' ) OR Colour LIKE '%" & Job & "%')) AND " & WareHouses& " order by Colour ASC"
end if
Dim rs
Set rs = Server.CreateObject("adodb.recordset")
Set rs = GetDisconnectedRS(strSQL, DBConnection)

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill/Painted " & Job & " </li>"
if job = "" then
job = "ALL"
end if
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Mill</th><th>" & Job & ": Allocated </th><th>Painted: " & Job & "</th><th>Pending</th><th>Other Inventory</th><th>Min level</th><th>Alerts</th></tr>"

rs2.movefirst
	do while not rs2.eof
	partqty = 0
	partqty2 = 0
	partqty3 = 0
	partqty4 = 0
	allocatedqty = 0
	
'Set rs = Server.CreateObject("adodb.recordset")
if Job = "" or Job = "ALL" then
	strSQL = "SELECT * FROM Y_INV WHERE " & Warehouses & " And Part = '" & rs2("Part") & "' order by Colour ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE ((Colour = 'Mill' AND Allocation Like '%" & Job &"%' ) OR Colour LIKE '%" & Job & "%')  And Part = '" & rs2("Part") & "' AND " & WareHouses & " order by Colour ASC"
end if
'rs.Cursortype = GetDBCursorType
'rs.Locktype = GetDBLockType
'rs.Open strSQL, DBConnection	

	rs.filter = "Part = '" & rs2("Part") & "'" 

	if rs.eof then
	else
	rs.movefirst
	do while not rs.eof
		IF RS("WAREHOUSE") = "GOREWAY" OR RS("Warehouse") = "NASHUA" OR RS("Warehouse") = "NPREP" OR RS("Warehouse") = "JUPITER" then
				if rs("colour") = "Mill" then
					if rs("allocation") = "" then
						partqty = rs("Qty") + partqty
					else
						allocatedqty = rs("Qty") + allocatedqty
					end if
				else
					partqty2 = rs("Qty") + partqty2
				end if
		End if
		IF RS("WAREHOUSE") = "DURAPAINT" OR RS("WAREHOUSE") = "HORNER" then
			partqty4 = rs("Qty") + partqty4
		end if
		

		IF rs("WAREHOUSE") = "SAPA" OR rs("WAREHOUSE") = "HYDRO" OR RS("WAREHOUSE") = "CAN-ART" OR RS("WAREHOUSE") = "EXTAL SEA" OR RS("WAREHOUSE") = "DEPENDABLE" OR RS("WAREHOUSE") = "Durapaint(WIP)" then
				partqty3 = rs("Qty") + partqty3

		End if

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if partqty + partqty3 < rs2("MinLevel") AND partqty2 + partqty3 < rs2("MinLevel") then
		MinLevelAlert = "Below Minimum"
	end if

	if job = "ALL" then
		if partqty2 = 0 and partqty = 0 and allocatedqty = 0 then
		else
			response.write "<tr><td><a href='stockbydie.asp?part=" & rs2("Part") & "&ticket=level' target='_self'>" & rs2("part") & "</a></td><td>" & rs2("description") & "</td><td>" & partqty & "</td><td>" & allocatedqty & "</td><td> " & partqty2 & "</td><td> " & partqty3 & "</td><td> " & partqty4 & "</td><td> " & rs2("MinLevel") & "</td><td> " & MinLevelAlert & "</td></tr>"
		end if 
	else
		if partqty2 = 0 and allocatedqty = 0 then
		else
			response.write "<tr><td><a href='stockbydie.asp?part=" & rs2("Part") & "&JOB=" & job & "&ticket=level' target='_self'>" & rs2("part") & "</a></td><td>" & rs2("description") & "</td><td>" & partqty & "</td><td>" & allocatedqty & "</td><td> " & partqty2 & "</td><td> " & partqty3 & "</td><td> " & partqty4 & "</td><td> " & rs2("MinLevel") & "</td><td> " & MinLevelAlert & "</td></tr>"
		end if 
	end if
	
'rs.close
'set rs = nothing
end if
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

