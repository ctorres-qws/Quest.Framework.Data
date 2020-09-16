<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stocktoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--Shows all items in Inventory with Today as the entry date in Goreway, Durapaint, and Sapa-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->
<!-- January 2015, Michael Bernholtz fixed error in the Bundle viewing - Bundle Variable was not declared in the Durapaint Section-->
<!-- Added Nashua Januuary 2017 -->
<!-- Added USA View February 2019 -->


<!-- STOCK TODAY E-MAIL generated using this page (in Table Form, if this page changes - edit the E-mail StockTodayEmail.asp, July 28th 2014 (NOT IN TEXAS YET)-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Today</title>
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
		currentDate = Request.Querystring("CDay")
		CDay = currentDate  
	if currentDate = "" then
		currentDate = Date()
	End if
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM Y_INV WHERE DATEIN = #" & currentDate & "# Order BY WAREHOUSE, PART")

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER order BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection



%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Stock Today</h1>
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
         
        <ul id="Profiles" title="Stock Entered <% response.write CurrentDate %>" selected="true">
         <li class="group"><a href="stocktodaytable.asp?part=<%response.write part%>&CDay=<%response.write CDay%>" target="_self" >Stock Today (Row Form) - Switch to Table Form</a></li>
 
 
 <%
 if CountryLocation = "USA" then
 
		
		rs.filter = "WAREHOUSE='Jupiter'"

		response.write "<li class='group'> --- Jupiter --- </li>"

		do while not rs.eof
		part = rs("part")
			rs2.filter = "Part = '" & part & "'"
			if rs2.eof then 
				Description = "N/A"
			else
				Description = rs2("Description")
			end if

		qty = rs("qty")
		id = rs("ID")
		Lft = rs("Lft")
		colour = rs("colour")
		PO = rs("po") 
		Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop

 else ' CANADA

rs.filter = "WAREHOUSE='GOREWAY'"

response.write "<li class='group'> --- GOREWAY --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='HORNER'"

response.write "<li class='group'> --- HORNER --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='NASHUA'"

response.write "<li class='group'> --- NASHUA --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='NPREP'"

response.write "<li class='group'> --- NASHUA PREP --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop


rs.filter = "WAREHOUSE= 'MILVAN'"

response.write "<li class='group'> --- MILVAN --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")


		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft  %></a></li>
<%

rs.movenext
loop


rs.filter = "WAREHOUSE='TILTON'"

response.write "<li class='group'> --- TILTON --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'Allocation added for Pending
Allocation = rs("Allocation")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<%

rs.movenext
loop

rs.filter = "WAREHOUSE='SAPA' or WAREHOUSE='HYDRO'"

response.write "<li class='group'> --- HYDRO PENDING --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'Allocation added for Pending
Allocation = rs("Allocation")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<%

rs.movenext
loop

rs.filter = "WAREHOUSE='DURAPAINT'"

response.write "<li class='group'> --- DURAPAINT PENDING --- </li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'No Allocation for Durapaint
		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft  & "' " %></a></li>
<%

rs.movenext
loop

rs.filter = "WAREHOUSE='DURAPAINT(WIP)'"

response.write "<li class='group'> --- DURAPAINT (WIP) PENDING --- </li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'No Allocation for Durapaint
		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft  & "' " %></a></li>
<%

rs.movenext
loop


rs.filter = "WAREHOUSE='DEPENDABLE'"

response.write "<li class='group'> --- DEPENDABLE PENDING --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'Allocation special for Dependable
Allocation = rs("Allocation")

		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<%

rs.movenext
loop


rs.filter = "WAREHOUSE='METRA'"

response.write "<li class='group'> --- METRA PENDING --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
'Allocation for Metra
Allocation = rs("Allocation")

		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<%

rs.movenext
loop

rs.filter = "WAREHOUSE='APEL'"

response.write "<li class='group'> --- APEL --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")

%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' " %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='TORBRAM'"

response.write "<li class='group'> --- TORBRAM --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
	'Allocation special for Dependable
Allocation = rs("Allocation")	

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft & "' - Allocated to: " & Allocation %></a></li>
<%

rs.movenext
loop



rs.filter = "WAREHOUSE= 'TILTON'"

response.write "<li class='group'> --- TILTON --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")


		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft  %></a></li>
<%

rs.movenext
loop


rs.filter = "WAREHOUSE='EXTAL SEA'"

response.write "<li class='group'> --- Extal PENDING --- </li>"

do while not rs.eof
part = rs("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")


%>

<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft %></a></li>
<%

rs.movenext
loop


rs.filter = "WAREHOUSE='EXTRUDEX'"

response.write "<li class='group'> --- EXTRUDEX PENDING --- </li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po") 
Bundle = rs("Bundle")
Allocation = rs("Allocation")
		

%>
<li><a href="stockbyrackedit.asp?ticket=intoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & Description & ", " & qty & " SL" & ", " & Colour & " " & PO & " / " & Bundle & " " & Lft  & "' " %></a></li>
<%

rs.movenext
loop

 end if ' USA/CANADA

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
Set DBConnection = nothing

%>
 </ul>                 
            
          
</body>
</html>
