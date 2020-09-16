<!--#include file="dbpath.asp"-->
    <!-- Updated May 9th to include length in feet, Michael Bernholtz -->
	<!-- Special Code to Overwrite Metra Door Material which ALWAYS comes in at 21.33 feet May 2018-->
	<!-- USA Seperation added February 2019-->
	<!-- Force Thickness of Gauge- June 2019-->
<!--Date: January 16, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add EXTRUDEX supplier
-->	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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

InventoryType = Request.Querystring("InventoryType")
If InventoryType= "" Then
	InventoryType = "Extrusion"
End If

currentDate = Date()

part = REQUEST.QueryString("part")

color = REQUEST.QueryString("color")
length = REQUEST.QueryString("length")
If length = "" Then
	length = 0
End If

lft = 0
lmm = 0
linch = 0
' Remove any length texts to ensure that length is a number
length = Replace(length, "'", "")
length = Replace(length, """", "")
length = Replace(length, "mm", "")
length = CINT(Length)

If length < 100 Then  'changed to 30 from 100 April 2016, shaun levy, returned to 100 Jan 2017
	linch = Round(length * 12,0)
	lmm = Round(linch * 25.4,0)
	lft = length
Else
	linch = length
	lmm = Round(linch * 25.4,0)
	lft = Round(length /12,0)
End If

If length > 300 Then
	linch = Round(length / 25.4,0)
	lmm = length
	lft = Round(length / 304.8,0)
End If

'Special Code to Overwrite Metra Door Material which ALWAYS comes in at 21.33 feet
'Check for NC beginning and then Overwrite Length entered
' Specific instruction by Shaun Levy and Gunja Bhatt

if UCASE(Left(part,2)) = "NC" then
	linch = 256
	lmm = 6502.4
	lft = 21.33
end if

aisle = REQUEST.QueryString("aisle")

if Len(aisle) = 1 then
	aisle = Left(UCASE(aisle),1)
end if
If aisle = "I" Then
	aisle = "i"
End If
	if Len(aisle) = 2 then
		aisle = Left(UCASE(aisle),1) & Right(LCase(aisle),1)
	end if
If aisle = "in" or aisle = "In" or aisle = "IN" or aisle = "Inside" or aisle = "INSIDE" Then
	aisle = "inside"
End If
If aisle = "out" or aisle = "Out" or aisle = "OUT" or aisle = "Outside" or aisle = "OUTSIDE" Then
	aisle = "outside"
End If
warehouse = REQUEST.QueryString("warehouse")
po = REQUEST.QueryString("PO")
colorpo = REQUEST.QueryString("ColorPO")
bundle = REQUEST.QueryString("Bundle")
exbundle = REQUEST.QueryString("ExBundle")
allocation = REQUEST.QueryString("Allocation")
rack = REQUEST.QueryString("rack")
shelf = REQUEST.QueryString("shelf")
Thickness = REQUEST.QueryString("Thickness")
'Force Thickness - June 2019
if Part ="16_Gauge_Galv" then
	Thickness = 0.05
end if
if Part ="24_Gauge_Galv" then
	Thickness = 0.027
end if

qty = REQUEST.QueryString("qty")
If qty = "" Then
	qty = 0
End If
Width = REQUEST.QueryString("Width")
If Width = "" Then
	Width = 0
End If
height = REQUEST.QueryString("Height")
If Height = "" Then
	Height = 0
End If

allocation = REQUEST.QueryString("allocation")
expdate = REQUEST.QueryString("expdate")

If isDate(expdate) Then
	expDay = day(expdate)
	expMonth = month(expdate)
	expYear = year(expdate)
Else
	expdate = NULL
End If
	
fail = "0"

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = GetDBCursorTypeInsert
rs.Locktype = GetDBLockTypeInsert
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_INVLOG"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_MASTER"
rs3.Cursortype = GetDBCursorTypeInsert
rs3.Locktype = GetDBLockTypeInsert
rs3.Open strSQL, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_COLOR"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL, DBConnection

' part = REQUEST.QueryString("part") Moved up to top May 2018

PARTEXISTS = "NO"
rs3.filter = "PART ='" & TRIM(Part) & "'"
if rs3.eof then 
	PARTEXISTS = "NO"
else
	PARTEXISTS = "YES"
end if
IF PARTEXISTS = "NO" THEN

	RS3.ADDNEW
	RS3.FIELDS("PART") = PART
	If GetID(isSQLServer,2) <> "" Then RS3.Fields("ID") = GetID(isSQLServer,2)
	RS3.UPDATE
	Call StoreID2(isSQLServer, RS3.Fields("ID"))
	rs3.movelast
	pid = rs3("ID")

END IF

rs.movelast

if rs.Fields("Part") = part AND rs.Fields("QTY") = qty AND rs.Fields("colour") = color AND rs.Fields("allocation") = allocation AND rs.Fields("warehouse") = warehouse and rs.Fields("PO") = PO and rs.Fields("Bundle") = Bundle then
fail = "1"
else

	rs.AddNew
	rs.Fields("Part") = part
	rs.Fields("colour") = color
	rs.Fields("qty") = qty
	rs.Fields("width") = width
	rs.Fields("height") = height
	rs.Fields("firstqty") = qty
	rs.Fields("linch") = linch
	rs.Fields("Thickness") = Thickness
	rs.Fields("lmm") = lmm
	rs.Fields("lft") = lft
	rs.Fields("warehouse") = warehouse
	
	if (warehouse = "HYDRO" or warehouse = "METRA" or warehouse = "APEL" or warehouse = "KEYMARK" or warehouse = "CAN-ART" or warehouse = "EXTRUDEX") then
		rs.Fields("supplier") = warehouse
	end if
	if warehouse = "SAPA" or warehouse = "SAPAMILL" or warehouse = "SAPAMONTREAL" then
			rs.Fields("supplier") = "HYDRO"
	end if
	
	rs.Fields("PO") = PO
	rs.Fields("ColorPO") = colorPO
	rs.Fields("Bundle") = Bundle
	rs.Fields("EXBundle") = ExBundle
	rs.Fields("aisle") = trim(aisle)
	rs.Fields("rack") = trim(rack)
	rs.Fields("shelf") = trim(shelf)
	rs.Fields("allocation") = allocation
	'Temporary addition until Warehouses resolved October 2018
	if inventorytype = "Sheet" and warehouse = "NASHUA" then
		rs.Fields("LabelPrint") = "No"
	end if


rs.Fields("DateIn") = FixSQLDate(currentDate, isSQLServer)
rs.Fields("ModifyDate") = FixSQLDate(currentDate, isSQLServer)

	'code to create PREF (Gathers PREF name and COLOUR Code)

	rs3.filter = "PART = '" & part & "'" 
	if not rs3.eof then
		PREFName = RS3("PREFREF")
	else	
		PREFName = "x"
	end if

	rs3.filter = "" 
	rs4.filter = "PROJECT = '" & color & "'"
	if not rs4.eof then
		PREFColour = RS4("CODE")
	else
		PREFColour = "x"
	end if
	rs4.filter = "" 
	PREFValue = Trim(PREFName & " " & PREFColour)
	rs.Fields("PREF") = PREFValue

	' end of new code - June 2015 (except one line for Invlog later)

	if warehouse = "WINDOW PRODUCTION" or warehouse = "COM PRODUCTION" or warehouse = "JUPITER PRODUCTION" or warehouse = "SCRAP" then
		rs.Fields("DateOut") = currentDate
	end if

	if isDate(expdate) then
		rs.Fields("ExpectedDate") = expdate 'FixSQLDate(expdate, isSQLServer)
	end if

	If UCase(Request("PrefImport")) = "TRUE" Then
		rs.Fields("Note 3") = "PrefImport: Auto Import"
		rs.Fields("Note 4") = Request("PrefID")
		If UCase(InventoryType) = "SHEET" Then
			rs.Fields("Height") = Request("Height")
			rs.Fields("Width") = Request("Width")
		End If
	End If

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	rs.update

	Call StoreID1(isSQLServer, rs.Fields("ID"))
	itemid=  rs.fields("ID")

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

	rs2.AddNew
	rs2.Fields("Part") = part
	rs2.Fields("colour") = color
	rs2.Fields("qty") = qty
	rs2.Fields("width") = width
	rs2.Fields("height") = height
	rs2.Fields("firstqty") = qty
	rs2.Fields("linch") = linch
	rs2.Fields("Thickness") = Thickness
	rs2.Fields("lmm") = lmm
	rs2.Fields("lft") = lft
	rs2.Fields("aisle") = trim(aisle)
	rs2.Fields("rack") = trim(rack)
	rs2.Fields("shelf") = trim(shelf)
	rs2.Fields("allocation") = allocation
	rs2.Fields("warehouse") = warehouse
	rs2.Fields("PO") = po
	rs2.Fields("colorPO") = colorpo
	rs2.Fields("Bundle") = Bundle
	rs2.Fields("ExBundle") = ExBundle
	rs2.Fields("transaction") = "enter"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ModifyDate") = FixSQLDate(currentDate, isSQLServer)
	rs2.Fields("ItemId") = itemid
	rs2.Fields("PREF") = PREFValue
	if warehouse = "WINDOW PRODUCTION" or warehouse = "COM PRODUCTION" or warehouse = "JUPITER PRODUCTION" or warehouse = "SCRAP" then
		rs2.Fields("DateOut") = currentDate
	end if

	if isDate(expdate) then
		rs2.Fields("ExpectedDate") = expdate 'FixSQLDate(expdate, isSQLServer)
	end if

	rs2.update

end if

DbCloseAll

End Function

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Stock Input</h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stock.asp#_enter" target="_self">Enter Stock</a>
    </div>


    
<ul id="Report" title="Stock Entered" selected="true">
<% if fail = "1" then
response.write "<li>Duplicate Inventory Item not entered</li>"
else
%>
	<li><% response.write "Part " & part %></li>
	<li><% response.write "Color " & color %></li>
	<li><% response.write "Allocated to: " & allocation %></li>
    <li><% response.write "Qty " & Qty %></li>
    <li><% response.write "Length " & Round(linch,2) & "''" %></li>
    <li><% response.write "Warehouse " & warehouse %></li>
    <li><% response.write "Aisle " & aisle %></li>
    <li><% response.write "Rack " & rack %></li>
    <li><% response.write "Shelf " & shelf %></li>
    <li><% response.write "PO " & PO %></li>
	<li><% response.write "Color PO " & ColorPO %></li>
	<li><% response.write "Bundle " & Bundle %></li>
	<li><% response.write "Ext. Bundle " & ExBundle %></li>
	<li><% response.write "Allocation: " & Allocation %></li>
	
	
	<% 
	if isDate(expdate) then
		response.write "<li>Expected Date " & expdate & "</li>" 
	else
		response.write "No Date Entered / or Date was in wrong format" 
	end if
	%>
	
    <% if PARTEXISTS = "NO" then %>
    	<li class="group">Added to Master</li>
        <li><a href="MASTEReditform.asp?id=<% response.Write pid %>&part=<% response.write part %>" target="_self">Part Added, Need Kg/m</a></li>
    <% end if %>
<%
end if
		response.write "<li><a class = 'whiteButton' href='Stock.asp?InventoryType=" & InventoryType & "#_enter' target='_self'>Add Another Item </a></li>"

if CountryLocation = "USA" then
	response.write "<li><a class = 'whiteButton' href='IndexTexas.html#_Inv' target='_self'>Home</a></li>"
else
	response.write "<li><a class = 'whiteButton' href='Index.html#_Inv' target='_self'>Home</a></li>"
end if
		
	%>	
</ul>

<%

%>

</body>
</html>

<%

'rs.close
'set rs=nothing
'rs2.close
'set rs2=nothing
'rs3.close
'set rs3 = nothing
'rs4.close
'set rs4 = nothing
'DBConnection.close
'set DBConnection=nothing

%>
