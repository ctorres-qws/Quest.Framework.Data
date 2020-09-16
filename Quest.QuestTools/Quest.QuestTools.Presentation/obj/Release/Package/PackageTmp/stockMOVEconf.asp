<!--#include file="dbpath.asp"-->
    <!-- Updated May 9th to include Length in Feet, Michael Bernhotlz -->
	<!-- USA included - February 2019 - Michael Bernholtz -->
	<!-- updated to include Datein field - unsure why it is was not there December 2020 - michael Bernholtz-->
					
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

ticket = Request.Querystring("ticket")
part = REQUEST.QueryString("part")
pid = request.querystring("id")
aisle = request.querystring("aisle")
FloorNote = request.querystring("FloorNote")
FloorNote2 = request.querystring("FloorNote2")
inventoryType = request.querystring("InventoryType")
Supplier = request.querystring("Supplier")
			if isnull(Supplier) then
				Supplier = ""
			End if

poSEARCH = request.querystring("poSEARCH")
bundleSEARCH = request.querystring("bundleSEARCH")
pobundleSEARCH = request.querystring("pobundleSEARCH")

	width = REQUEST.QueryString("width")
	if width = "" then
		width = 0
	end if
	height = REQUEST.QueryString("height")
	if height = "" then
		height = 0
	end if

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


colour = REQUEST.QueryString("colour")
qty = REQUEST.QueryString("qty")

QtyMove = Request.QueryString("QtyMove")
BundleMove = request.querystring("BundleMove")
ExBundleMove = request.querystring("ExBundleMove")
WarehouseMove = Request.Querystring("WarehouseMove")
UpdateSuccess = False

allocation = REQUEST.QueryString("allocation")
length = REQUEST.QueryString("length")
if length = "" then
	length = 0
end if

currentDate = Date()

if length < 100 then
linch = Round(length * 12,0)
lmm = Round(linch * 25.4,0)
lft = length
else
linch = length
lmm = Round(linch * 25.4,0)
lft = Round(length /12,0)
end if

if length > 300 then
linch = Round(length / 25.4,0)
lmm = length
lft = Round(length / 304.8,0)
end if

'Special Code to Overwrite Metra Door Material which ALWAYS comes in at 21.33 feet
'Check for NC beginning and then Overwrite Length entered
' Specific instruction by Shaun Levy and Gunja Bhatt

if UCASE(Left(part,2)) = "NC" then
	linch = 256
	lmm = 6502.4
	lft = 21.33
end if

color = REQUEST.QueryString("color")
aisle = REQUEST.QueryString("aisle")

if Len(aisle) = 1 then
	aisle = Left(UCASE(aisle),1)
end if
if aisle = "I" then
	aisle = "i"
end if
	if Len(aisle) = 2 then
		aisle = Left(UCASE(aisle),1) & Right(LCase(aisle),1)
	end if
rack = REQUEST.QueryString("rack")
shelf = REQUEST.QueryString("shelf")
po = REQUEST.QueryString("PO")
colorpo = REQUEST.QueryString("ColorPO")
bundle = REQUEST.QueryString("bundle")
exbundle = REQUEST.QueryString("exbundle")
thickness = REQUEST.QueryString("thickness")
if thickness = "" then
	thickness = 0
end if
warehouse = REQUEST.QueryString("warehouse")
expdate = ""
expdate = request.querystring("expdate")
JobComplete = request.querystring("JobComplete")
JobComplete1 = request.querystring("JobComplete1")
'read from the Move version, so JobComplete1 not JobComplete like the edit
colorpo1 = REQUEST.QueryString("ColorPO1")
LengthFt = REQUEST.QueryString("LengthFt")
DateOut = ""

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
	If isSQLServer Then strSQL = "SELECT * FROM Y_INV WHERE ID=" & pid
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection
	rs.filter = "ID = " & pid

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * From Y_INVLOG"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL, DBConnection

' Marker to Create DateOut when Warehouse Becomes Production for the first time
'July 15th, 2014 - Michael Bernholtz at Request of Shaun Levy
	If UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION" or UCASE(warehouse) = "JUPITER PRODUCTION" or UCASE(warehouse) = "SCRAP" Then
		If UCASE(rs.Fields("warehouse")) <> "WINDOW PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "COM PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "JUPITER PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "SCRAP" Then
			DateOut = currentDate
		End If
	End If

	If (UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION" or UCASE(warehouse) = "JUPITER PRODUCTION") and (UCASE(warehouseMove) = "WINDOW PRODUCTION" or UCASE(warehouseMove) = "COM PRODUCTION" or UCASE(warehouseMove) = "JUPITER PRODUCTION") Then
		UpdateSuccess = FALSE
		ErrorReason = "PROD"
'CODE added May 2016 to prevent Production items from being transfered within production.
'This type of transfer will alter daily production values and has to be cleared
	Else
		If QTY- QTYMOVE >=1 Then
			UpdateSuccess = TRUE

' Add to Y_InvLog
			rs2.AddNew
			rs2.Fields("Part") = rs.Fields("Part")
			rs2.Fields("colour") = rs.Fields("colour")
			rs2.Fields("qty") = rs.Fields("qty")
			rs2.Fields("width") = rs.Fields("width")
			rs2.Fields("height") = rs.Fields("height")
			rs2.Fields("linch") = rs.Fields("linch")
			rs2.Fields("lmm") = rs.Fields("lmm")
			rs2.Fields("lft") = rs.Fields("lft")
			rs2.Fields("warehouse") = rs.Fields("warehouse")
			rs2.Fields("aisle") = rs.Fields("aisle")
			rs2.Fields("rack") = rs.Fields("rack")
			rs2.Fields("shelf") = rs.Fields("shelf")
			rs2.Fields("po") = rs.Fields("po")
			rs2.Fields("allocation") = rs.Fields("allocation")
			rs2.Fields("colorpo") = rs.Fields("colorpo")
			rs2.Fields("bundle") = BundleMove
			rs2.Fields("ExBundle") = ExBundleMove
			rs2.Fields("PREF") = rs.Fields("PREF")
			rs2.Fields("thickness") = rs.Fields("thickness")
			rs2.Fields("datein") = rs.Fields("datein")
			rs2.Fields("transaction") = "original"
			rs2.Fields("day") = cday
			rs2.Fields("month") = cmonth
			rs2.Fields("year") = cyear
			rs2.Fields("week") = weeknumber
			rs2.Fields("time") = cctime
			rs2.Fields("ModifyDate") = currentDate
			rs2.fields("Note") = rs.Fields("Note")
			rs2.fields("JobComplete") = rs.Fields("JobComplete")
			rs2.fields("Itemid") = pid
			If isDate(rs.Fields("ExpectedDate")) Then
				rs2.Fields("ExpectedDate") = rs.Fields("ExpectedDate")
			End If
			If isDate(rs.Fields("DateOut")) Then
				rs2.Fields("DateOut") = rs.Fields("DateOut")
			End If
			rs2.update

' Add to Y_InvLog
			rs2.AddNew
			rs2.Fields("Part") = part
			rs2.Fields("colour") = color
			rs2.Fields("qty") = qty - qtyMove
			rs2.Fields("width") = width
			rs2.Fields("height") = height
			rs2.Fields("linch") = linch
			rs2.Fields("lmm") = lmm
			rs2.fields("lft") = lft
			rs2.Fields("warehouse") = warehouse
			rs2.Fields("aisle") = trim(aisle)
			rs2.Fields("rack") = trim(rack)
			rs2.Fields("shelf") = trim(shelf)
			rs2.Fields("po") = po
			rs2.Fields("colorpo") = colorpo
			rs2.Fields("bundle") = bundle
			rs2.Fields("exbundle") = exbundle
			rs2.Fields("thickness") = thickness
			If UCASE(allocation) = "UNALLOCATED" Then
			Else
				rs2.Fields("Allocation") = allocation
			End If
			rs2.Fields("transaction") = "edit"
			rs2.Fields("day") = cday
			rs2.Fields("month") = cmonth
			rs2.Fields("year") = cyear
			rs2.Fields("week") = weeknumber
			rs2.Fields("time") = cctime
			rs2.Fields("ModifyDate") = currentDate
			rs2.fields("Note") = FloorNote
			rs2.fields("JobComplete") = JobComplete
			If isDate("ExpectedDate") Then
				rs2.Fields("ExpectedDate") = expdate
			End If
			rs2.Fields("datein") = rs.Fields("datein")
			If isDate(DateOut) Then
				rs2.Fields("DateOut") = DateOut
			End If
			rs2.update

			rs.Fields("Part") = part
			rs.Fields("colour") = color
			rs.Fields("qty") = qty - QtyMove
			rs.Fields("width") = width
			rs.Fields("height") = height
			rs.Fields("linch") = linch
			rs.Fields("lmm") = lmm
			rs.fields("lft") = lft
			rs.Fields("warehouse") = warehouse
			rs.Fields("aisle") = trim(aisle)
			rs.Fields("rack") = trim(rack)
			rs.Fields("shelf") = trim(shelf)
			rs.Fields("po") = po
			rs.Fields("colorpo") = colorpo
			rs.Fields("bundle") = bundle
			rs.Fields("exbundle") = exbundle
			rs.Fields("thickness") = thickness
			If UCASE(allocation) = "UNALLOCATED" Then
			Else
				rs.Fields("Allocation") = allocation
			End If
			rs.Fields("ModifyDate") = currentDate
			rs.fields("Note") = FloorNote
			rs.fields("JobComplete") = JobComplete
			If isDate(expdate) Then
				rs.Fields("ExpectedDate") = expdate
			End If
			If isDate(DateOut) Then
				rs.Fields("DateOut") = DateOut
			End If
			rs.update

' -----------------------------------ADD NEW RECORD-------------------------------------------	

			rs.AddNew
			rs.Fields("Part") = part
			rs.Fields("colour") = color
			rs.Fields("qty") = qtyMove
			rs.Fields("width") = width
			rs.Fields("height") = height
			rs.Fields("linch") = linch
			rs.Fields("lmm") = lmm
			rs.Fields("lft") = lft
			rs.Fields("warehouse") = warehouseMOVE
			rs.Fields("PO") = PO
			rs.Fields("ColorPO") = colorpo1
			rs.Fields("bundle") = BundleMove
			rs.Fields("Supplier") = Supplier

			If UCase(warehouseMOVE) = "NPREP" Then
				rs.Fields("aisle") = "Zf1"
				rs.Fields("rack") = ""
				rs.Fields("shelf") = ""
			Else
				rs.Fields("aisle") = trim(aisle)
				rs.Fields("rack") = trim(rack)
				rs.Fields("shelf") = trim(shelf)
			End If

			rs.Fields("DateIn") = currentDate
			rs.Fields("exbundle") = exbundleMove
			rs.Fields("thickness") = thickness
			If JobComplete1 = "Unallocated" Then
			Else
				rs.Fields("Allocation") = JobComplete1
			End If
			rs.Fields("ModifyDate") = currentDate
			rs.fields("Note") = FloorNote2
			If JobComplete1 = "Unallocated" Then
			Else
				rs.fields("JobComplete") = JobComplete1
			End If


			If UCASE(warehouseMOVE) = "WINDOW PRODUCTION" or UCASE(warehouseMOVE) = "COM PRODUCTION" or UCASE(warehouseMOVE) = "JUPITER PRODUCTION" or UCASE(warehouseMOVE) = "SCRAP" Then
				rs.Fields("DateOut") = currentDate
			End If
			
			If UCASE(warehouseMOVE) = "NASHUA" AND inventoryType = "Sheet" Then
			'If UCASE(warehouseMOVE) = "NASHUA" AND (UCASE(warehouse) = "SAPA" OR UCASE(warehouse) = "HYDRO" ) AND inventoryType = "Sheet" Then
						
			rs.Fields("LabelPrint") = "No"
			End If

			If isDate(expdate) Then
				rs.Fields("ExpectedDate") = expdate
			End If
			If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
			rs.update
			Call StoreID1(isSQLServer, rs.Fields("ID"))

' Add to Y_InvLog
			rs2.AddNew
			rs2.Fields("Part") = part
			rs2.Fields("colour") = color
			rs2.Fields("qty") = qtyMove
			rs2.Fields("width") = width
			rs2.Fields("height") = height
			rs2.Fields("linch") = linch
			rs2.Fields("lmm") = lmm
			rs2.Fields("lft") = lft

			If UCase(warehouseMOVE) = "NPREP" Then
				rs2.Fields("aisle") = "Zf1"
				rs2.Fields("rack") = ""
				rs2.Fields("shelf") = ""
			Else
				rs2.Fields("aisle") = trim(aisle)
				rs2.Fields("rack") = trim(rack)
				rs2.Fields("shelf") = trim(shelf)
			End If

			rs2.Fields("warehouse") = warehouseMove
			rs2.Fields("PO") = po
			rs2.Fields("colorPO") = colorpo1
			rs2.Fields("Bundle") = Bundle
			rs2.Fields("ExBundle") = ExBundle
			rs2.Fields("transaction") = "transfer"
			rs2.Fields("day") = cday
			rs2.Fields("month") = cmonth
			rs2.Fields("year") = cyear
			rs2.Fields("week") = weeknumber
			rs2.Fields("time") = cctime
			rs2.Fields("ModifyDate") = currentDate
			rs2.Fields("thickness") = thickness
			If JobComplete1 = "Unallocated" Then
			Else
				rs2.Fields("Allocation") = JobComplete1
			End If
			rs2.Fields("itemid") = pid
			rs2.fields("Note") = FloorNote
			If JobComplete1 = "Unallocated" Then
			Else
				rs2.fields("JobComplete") = JobComplete1
			End If
			rs2.Fields("datein") = rs.Fields("datein")
			
			If UCASE(warehouseMOVE) = "WINDOW PRODUCTION" or UCASE(warehouseMOVE) = "COM PRODUCTION"  or UCASE(warehouseMOVE) = "JUPITER PRODUCTION" or UCASE(warehouseMOVE) = "SCRAP" Then
				rs2.Fields("DateOut") = currentDate
			End If

			If isDate(expdate) Then
				rs2.Fields("ExpectedDate") = expdate
			End If

			rs2.update

			UpdateSuccess = True

			If isSQLServer Then
				Dim str_Referrer: str_Referrer = Request("ReferrerPage")
				Call PrefAddTransfer(GetID(isSQLServer,1), warehouse, warehouseMOVE, part, qtyMove, color, LengthFt, False, "StockMove", str_Note, po, str_Referrer)
			End If

		End If
	End If



'rs.close
'set rs=nothing
'rs2.close
'set rs2=nothing
'DBConnection.close
'set DBConnection=nothing

	DbCloseAll

End Function

%>

	</head>
<body>
<div id="clock"></div>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockbyrackedit.asp?id=<% response.write pid %>&aisle=<% response.write aisle %>&ticket=<% response.write ticket%>&pobundleSEARCH=<% response.write pobundleSEARCH%>&poSEARCH=<% response.write poSEARCH%>&bundleSEARCH=<% response.write bundleSEARCH%>&part=<% response.write part%>" target="_self">Edit Stock</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="stockbyrackedit.asp?id=<% response.write pid %>&aisle=<% response.write aisle %>&ticket=<% response.write ticket%>&pobundleSEARCH=<% response.write pobundleSEARCH%>&poSEARCH=<% response.write poSEARCH%>&bundleSEARCH=<% response.write bundleSEARCH%>&part=<% response.write part%>" method="GET" target="_self" selected="true" >              

  <%
  if UpdateSuccess = TRUE then
  %>
        <h2>Transfer Complete</h2>

        <BR>
       <p> Stock has been updated to: <%response.write QTY-QtyMOVE %> in: <%response.write warehouse %></p>
	   <p> A new stock record of:  <%response.write QtyMOVE %> has been added to: <%response.write warehouseMOVE %></p>
	   <br>
	   
		<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>">
		<input type="hidden" name='pobundleSearch' id='pobundleSearch' value="<%response.write pobundleSEARCH %>">
		<input type="hidden" name='bundleSearch' id='bundleSearch' value="<%response.write bundleSEARCH %>">
		<input type="hidden" name='poSearch' id='poSearch' value="<%response.write poSEARCH %>">
        <input type="hidden" name='part' id='part' value="<%response.write part %>">
        <input type="hidden" name='id' id='id' value="<%response.write pid %>">
         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
<%
	else
%>
	<h2>Could Not Transfer</h2>
<%
	if ErrorReason = "PROD" then
%>
	<p> Cannot Transfer Records from PRODUCTION to PRODUCTION </p>
<%
	else
%>
	<p> Quantity to Move (<%Response.Write QTYMOVE%>) must be less than Total Quantity of item(<%Response.write QTY%>) </p>
<%
	end if
%>
	
	<p>  </p>
	<BR>
	    <input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>">
		<input type="hidden" name='pobundleSearch' id='pobundleSearch' value="<%response.write pobundleSEARCH %>">
		<input type="hidden" name='bundleSearch' id='bundleSearch' value="<%response.write bundleSEARCH %>">
		<input type="hidden" name='poSearch' id='poSearch' value="<%response.write poSEARCH %>">
        <input type="hidden" name='part' id='part' value="<%response.write part %>">
        <input type="hidden" name='id' id='id' value="<%response.write pid %>">
	 <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>

<%
	end if
%>

            </form>

</body>
</html>



