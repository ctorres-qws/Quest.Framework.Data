
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 <!-- February 2019 - Added USA view -->
<!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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
ticket = request.QueryString("ticket")		
filterBundle = request.QueryString("poBundle")	
if filterBundle = "" then
	filterBundle= request.QueryString("poBundleSEARCH")
end if

if CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE = 'JUPITER' or WAREHOUSE = 'JUPITER PRODUCTION') AND (PO LIKE '%" & filterbundle & "%' OR BUNDLE LIKE '%" & filterbundle & "%' OR EXBUNDLE LIKE '%" & filterbundle & "%' OR COLORPO LIKE '%" & filterbundle & "%') ORDER BY WAREHOUSE, PART, DATEIN DESC"
else
	strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE <> 'JUPITER' OR WAREHOUSE <> 'JUPITER PRODUCTION') AND(PO LIKE '%" & filterbundle & "%' OR BUNDLE LIKE '%" & filterbundle & "%' OR EXBUNDLE LIKE '%" & filterbundle & "%' OR COLORPO LIKE '%" & filterbundle & "%') ORDER BY WAREHOUSE, PART, DATEIN DESC"
end if



Set rs = Server.CreateObject("adodb.recordset")

'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>

<%
	Select Case ticket
		Case "old"
%>
		<a class="button leftButton" type="cancel" href="stockoldreport.asp" target="_self">Old Stock</a>
<%
		Case else
		%>
		<a class="button leftButton" type="cancel" href="allbypobundle1.asp" target="_self">PO / Bundle</a>
<%
	End Select
%>

        </div>

        <ul id="Profiles" title="Profiles" selected="true">

<% 

'Loops through all Warehouses and then posts for each one.

if CountryLocation = "USA" then
	strSQL2 = "SELECT * FROM Y_WAREHOUSE WHERE Country ='USA' ORDER BY ID ASC"
else
	strSQL2 = "SELECT * FROM Y_WAREHOUSE WHERE Country ='Canada' ORDER BY ID ASC"
end if


Set rs2 = Server.CreateObject("adodb.recordset")
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

rs2.movefirst
Do While Not rs2.eof
	WarehouseName = rs2("Name")

	' Each Warehouse
	rs.filter = "Warehouse = '" & WarehouseName & "'" 

	response.write "<li class='group'>" & WarehouseName & "</li>"
	do while not rs.eof
%>
		<li><a href="stockbyrackedit.asp?ticket=allbp&POBundle=<%response.write filterBundle %>&id=<% response.write rs("ID") %>" target="_self"> 
		<%response.write rs("part") & ", " & rs("qty") & " SL" & ", " & rs("lft") & "', " & rs("colour") & " -PO: " & rs("po") 
		
		if rs("Allocation") = "" or isnull(rs("Allocation")) then
		else
		response.write "  | Allocation: " & rs("Allocation")
		end if
		
		if rs("Bundle") = "" or isnull(rs("Bundle")) then
		else
		response.write "  | Bundle: " & rs("Bundle")
		end if
		
		if rs("EXBundle") = "" or isnull(rs("EXBundle")) then
		else
		response.write "  | EX Bundle: " & rs("EXBundle")
		end if
		
		
		if rs("ExpectedDate") = "" or isnull(rs("ExpectedDate")) or (WarehouseName = "GOREWAY" or WarehouseName = "HORNER" or WarehouseName = "JUPITER") then
		else
		response.write " Expected: " & rs("ExpectedDate")
		end if
		
				 
		if (RS("Width") > 1 and RS("Height") > 1) then
		 response.write " Size:" & rs("Width") & " X " & rs("Height") 
		end if
		
		if rs("DateIn") = "" or isnull(rs("DateIn")) then
		else
			if WarehouseName = "WINDOW PRODUCTION" or WarehouseName = "COM PRODUCTION" or WarehouseName = "SCRAP" or WarehouseName = "JUPITER PRODUCTION" then
			response.write " | Exit: " & rs("DateOut")
			else
			response.write " | Entered: " & rs("DateIn")
			end if
		end if
		
		'addition of MILVAN by Annabel Ramirez Feb 13, 2020
		if (WarehouseName = "GOREWAY" or WarehouseName = "NASHUA" or WarehouseName = "JUPITER" or WarehouseName = "MILVAN") then
			if (RS("Aisle") = "" AND RS("Rack") = "" AND RS("Shelf") = "") then
				else
			response.write " -- " & rs("Aisle") & " - " & rs("Rack") & " - " & rs("Shelf")
				end if
		end if

		%>
		</a></li>
<%

	rs.movenext
	loop

rs2.movenext
loop

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
set DBConnection = nothing

%>
      </ul>
</body>
</html>
