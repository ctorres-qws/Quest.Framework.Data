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
	Dim str_Warehouse_Entry, str_Warehouse_Exit
	ticket = Request.Querystring("ticket")
	part = REQUEST.QueryString("part")
	pid = request.querystring("id")
	aisle = request.querystring("aisle")
	poSEARCH = request.querystring("poSEARCH")
	bundleSEARCH = request.querystring("bundleSEARCH")
	pobundleSEARCH = request.querystring("pobundleSEARCH")

	currentDate = Date()

	STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
	ccTime = hour(now) & ":" & minute(now)
	cDay = day(now)
	cMonth = month(now)
	cYear = year(now)
	currentDate = Date
	weekNumber = DatePart("ww", currentDate)

	colour = REQUEST.QueryString("colour")
	qty = REQUEST.QueryString("qty")
	qty_Orig = 0
	length = REQUEST.QueryString("length")
	if length ="" then
		length = 0
	end if

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
	if aisle = "in" or aisle = "In" or aisle = "IN" or aisle = "Inside" or aisle = "INSIDE" then
		aisle = "inside"
	end if
	if aisle = "out" or aisle = "Out" or aisle = "OUT" or aisle = "Outside" or aisle = "OUTSIDE" then
		aisle = "outside"
	end if
	
	

	rack = REQUEST.QueryString("rack")
	shelf = REQUEST.QueryString("shelf")
	po = REQUEST.QueryString("PO")
	colorpo = REQUEST.QueryString("ColorPO")
	bundle = REQUEST.QueryString("bundle")
	exbundle = REQUEST.QueryString("exbundle")
	allocation = REQUEST.QueryString("allocation")
	thickness = REQUEST.QueryString("thickness")
	FloorNote = ""
	FloorNote = REQUEST.QueryString("FloorNote")
	JobComplete = REQUEST.QueryString("JobComplete")
	StatusNote = REQUEST.QueryString("StatusNote")
	PREF = REQUEST.QueryString("pref")

	width = REQUEST.QueryString("width")
	if width = "" then
		width = 0
	end if
	height = REQUEST.QueryString("height")
	if height = "" then
		height = 0
	end if

	if thickness = "" then
		thickness = 0
	end if

	warehouse = REQUEST.QueryString("warehouse")
	expdate = ""
	expdate = request.querystring("expdate")

	LengthFt = REQUEST.QueryString("LengthFt")

' Marker to Create DateOut when Warehouse Becomes Production for the first time
'July 15th, 2014 - Michael Bernholtz at Request of Shaun Levy
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
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "ID = " & pid

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT top 1 * From Y_INVLOG"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL, DBConnection
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * From Y_MASTER"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL, DBConnection

	Set rs4 = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * From Y_COLOR"
	rs4.Cursortype = 2
	rs4.Locktype = 3
	rs4.Open strSQL, DBConnection
	
	' RS5 Discrepancy log opens in the IF statement

	if UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION" or UCASE(warehouse) = "JUPITER PRODUCTION" or UCASE(warehouse) = "SCRAP" then
		if UCASE(rs.Fields("warehouse")) <> "WINDOW PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "COM PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "JUPITER PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "SCRAP" then 
			DateOut = currentDate
		end if 
	end if

	str_Warehouse_Entry = warehouse
	str_Warehouse_Exit = rs.Fields("warehouse")

	qty_Orig = rs.Fields("qty")

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
	rs2.Fields("colorpo") = rs.Fields("colorpo")
	rs2.Fields("bundle") = rs.Fields("bundle")
	rs2.Fields("exbundle") = rs.Fields("exbundle")
	rs2.Fields("allocation") = rs.Fields("allocation")
	rs2.Fields("thickness") = rs.Fields("thickness")
	rs2.Fields("ItemId") = rs.Fields("ID")
	rs2.Fields("PREF") = rs.Fields("PREF")
	rs2.Fields("Note 2") = rs.Fields("Note 2")
	rs2.Fields("datein") = rs.Fields("datein")
	rs2.Fields("transaction") = "original"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ModifyDate") = currentDate
	rs2.Fields("Note") = rs.Fields("Note")
	rs2.Fields("JobComplete") = rs.Fields("JobComplete")
	if isDate(rs.Fields("ExpectedDate")) then
		rs2.Fields("ExpectedDate") = rs.Fields("ExpectedDate")
	end if
	if isDate(rs.Fields("DateOut")) then
		rs2.Fields("DateOut") = rs.Fields("DateOut")
	end if
	rs2.update

	rs2.AddNew
	rs2.Fields("Part") = part
	rs2.Fields("colour") = color
	rs2.Fields("qty") = qty
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
	if UCASE(allocation) = "UNALLOCATED" then
	else	
	rs2.Fields("Allocation") = allocation
	end if
	rs2.Fields("thickness") = thickness
	rs2.Fields("transaction") = "edit"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ModifyDate") = currentDate
	rs2.Fields("ItemId") = pid
	rs2.fields("PREF") = PREF
	rs2.fields("Note") = FloorNote
	rs2.fields("JobComplete") = JobComplete
	rs2.fields("Note 2") = StatusNote
	if isDate("ExpectedDate") then
		rs2.Fields("ExpectedDate") = expdate
	end if
	rs2.Fields("datein") = rs.Fields("datein")
	if isDate(DateOut) then
		rs2.Fields("DateOut") = DateOut
	end if
	rs2.update
	
	
	'Discrepancy Log Addition, Gets exact same records as Y_INV_LOG and then an additional record for Discrepancy
	' Discrepancy log only activated if Qty gets changed in EDIT
	if rs.fields("Qty") - qty = 0 then
	else
		Set rs5 = Server.CreateObject("adodb.recordset")
		strSQL5 = "SELECT * From Y_Discrepancy_Log "
		rs5.Cursortype = 2
		rs5.Locktype = 3
		rs5.Open strSQL5, DBConnection
	
		rs5.AddNew
		rs5.Fields("Part") = rs.Fields("Part")
		rs5.Fields("colour") = rs.Fields("colour")
		rs5.Fields("qty") = rs.Fields("qty")
		rs5.Fields("width") = rs.Fields("width")
		rs5.Fields("height") = rs.Fields("height")
		rs5.Fields("linch") = rs.Fields("linch")
		rs5.Fields("lmm") = rs.Fields("lmm")
		rs5.Fields("lft") = rs.Fields("lft")
		rs5.Fields("warehouse") = rs.Fields("warehouse")
		rs5.Fields("aisle") = rs.Fields("aisle")
		rs5.Fields("rack") = rs.Fields("rack")
		rs5.Fields("shelf") = rs.Fields("shelf")
		rs5.Fields("po") = rs.Fields("po")
		rs5.Fields("colorpo") = rs.Fields("colorpo")
		rs5.Fields("bundle") = rs.Fields("bundle")
		rs5.Fields("exbundle") = rs.Fields("exbundle")
		rs5.Fields("allocation") = rs.Fields("allocation")
		rs5.Fields("thickness") = rs.Fields("thickness")
		rs5.Fields("ItemId") = rs.Fields("ID")
		rs5.Fields("PREF") = rs.Fields("PREF")
		rs5.Fields("Note 2") = rs.Fields("Note 2")
		rs5.Fields("transaction") = "original"
		rs5.Fields("day") = cday
		rs5.Fields("month") = cmonth
		rs5.Fields("year") = cyear
		rs5.Fields("week") = weeknumber
		rs5.Fields("time") = cctime
		rs5.Fields("ModifyDate") = currentDate
		rs5.Fields("Note") = rs.Fields("Note")
		rs5.Fields("JobComplete") = rs.Fields("JobComplete")
		if isDate(rs.Fields("ExpectedDate")) then
			rs5.Fields("ExpectedDate") = rs.Fields("ExpectedDate")
		end if
		if isDate(rs.Fields("DateOut")) then
			rs5.Fields("DateOut") = rs.Fields("DateOut")
		end if
		rs5.update
	
		rs5.AddNew
		rs5.Fields("Part") = part
		rs5.Fields("colour") = color
		rs5.Fields("qty") = qty
		rs5.Fields("width") = width
		rs5.Fields("height") = height
		rs5.Fields("linch") = linch
		rs5.Fields("lmm") = lmm
		rs5.fields("lft") = lft
		rs5.Fields("warehouse") = warehouse
		rs5.Fields("aisle") = trim(aisle)
		rs5.Fields("rack") = trim(rack)
		rs5.Fields("shelf") = trim(shelf)
		rs5.Fields("po") = po
		rs5.Fields("colorpo") = colorpo
		rs5.Fields("bundle") = bundle
		rs5.Fields("exbundle") = exbundle
		if UCASE(allocation) = "UNALLOCATED" then
		else	
		rs5.Fields("Allocation") = allocation
		end if
		rs5.Fields("thickness") = thickness
		rs5.Fields("transaction") = "edit"
		rs5.Fields("day") = cday
		rs5.Fields("month") = cmonth
		rs5.Fields("year") = cyear
		rs5.Fields("week") = weeknumber
		rs5.Fields("time") = cctime
		rs5.Fields("ModifyDate") = currentDate
		rs5.Fields("ItemId") = pid
		rs5.Fields("PREF") = rs.Fields("PREF")
		rs5.fields("Note") = FloorNote
		rs5.fields("JobComplete") = JobComplete
		rs5.fields("Note 2") = StatusNote
		if isDate("ExpectedDate") then
			rs5.Fields("ExpectedDate") = expdate
		end if
		if isDate(DateOut) then
			rs5.Fields("DateOut") = DateOut
		end if
		rs5.update

		rs5.AddNew
		rs5.Fields("Part") = part
		rs5.Fields("colour") = color
		rs5.Fields("qty") = rs.fields("qty") - qty
		rs5.Fields("width") = rs.Fields("width")
		rs5.Fields("height") = rs.Fields("height")
		rs5.Fields("linch") = linch
		rs5.Fields("lmm") = lmm
		rs5.fields("lft") = lft
		rs5.Fields("warehouse") = warehouse
		rs5.Fields("aisle") = trim(aisle)
		rs5.Fields("rack") = trim(rack)
		rs5.Fields("shelf") = trim(shelf)
		rs5.Fields("po") = po
		rs5.Fields("colorpo") = colorpo
		rs5.Fields("bundle") = bundle
		rs5.Fields("exbundle") = exbundle
		if UCASE(allocation) = "UNALLOCATED" then
		else
		rs5.Fields("Allocation") = allocation
		end if
		rs5.Fields("thickness") = thickness
		rs5.Fields("transaction") = "Discrepancy"
		rs5.Fields("day") = cday
		rs5.Fields("month") = cmonth
		rs5.Fields("year") = cyear
		rs5.Fields("week") = weeknumber
		rs5.Fields("time") = cctime
		rs5.Fields("ModifyDate") = currentDate
		rs5.Fields("ItemId") = pid
		rs5.Fields("PREF") = rs.Fields("PREF")
		rs5.fields("Note") = FloorNote
		rs5.fields("JobComplete") = JobComplete
		rs5.fields("Note 7") = "SHOW"			' SHOW or HIDE (Hide will be set by Shaun in report)
		rs5.fields("Note 2") = StatusNote
		if isDate("ExpectedDate") then
			rs5.Fields("ExpectedDate") = expdate
		end if
		if isDate(DateOut) then
			rs5.Fields("DateOut") = DateOut
		end if
		rs5.update

	end if
	' End Discrepancy log Addition

	rs.Fields("Part") = part
	rs.Fields("colour") = color
	rs.Fields("qty") = qty
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
	if UCASE(allocation) = "UNALLOCATED" then
	else
		rs.Fields("Allocation") = allocation
	end if
	rs.Fields("thickness") = thickness
	rs.fields("Note") = FloorNote
	rs.fields("JobComplete") = JobComplete
	rs.fields("Note 2") = StatusNote
	
	rs.Fields("ModifyDate") = currentDate
	if isDate(expdate) then
		rs.Fields("ExpectedDate") = expdate
	end if
	if isDate(DateOut) then
		rs.Fields("DateOut") = DateOut
	end if
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
	PREFValue = PREFName & " " & PREFColour
	rs.Fields("PREF") = PREFValue

	' end of new code - June 2015 (except one line for Invlog later)

	rs.update

	If isSQLServer Then
		Dim str_Referrer: str_Referrer = Request("ReferrerPage")
		If CLng(qty_Orig) <> CLng(qty) Then

			If str_Warehouse_Exit <> str_Warehouse_Entry Then

				If CLng(qty_Orig) > CLng(qty) And CLng(qty_Orig) > 0 Then
					str_Note = "A1: qty_Orig(" & qty_Orig & ") > qty(" & qty & ")"
					Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
					Call PrefAddTransfer(pid, str_Warehouse_Exit, "ADJUSTMENT", part, qty_Orig - qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
				ElseIf CLng(qty) > CLng(qty_Orig) And CLng(qty_Orig) > 0 Then
					str_Note = "A2: qty(" & qty & ") > qty_Orig(" & qty_Orig & ")"
					'Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty_Orig, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
					'Call PrefAddTransfer(pid, "ADJUSTMENT", str_Warehouse_Entry, part, qty - qty_Orig, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
					Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
				Else
					str_Note = "A3: "
					Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)
				End If

			Else

				If str_Warehouse_Exit = str_Warehouse_Entry And UCase(str_Warehouse_Exit) = "NASHUA" Then

					If CLng(qty_Orig) > CLng(qty) And CLng(qty_Orig) > 0 Then
						str_Note = "B1: qty_Orig(" & qty_Orig & ") > qty(" & qty & ")"
						qty = qty_Orig - qty
						str_Warehouse_Entry = "NASHUA_ADJUSTMENT" ' Warehouse Exit
					ElseIf CLng(qty) > CLng(qty_Orig) And CLng(qty_Orig) > 0 Then
						str_Note = "B2: qty(" & qty & ") > qty_Orig(" & qty_Orig & ")"
						qty = qty - qty_Orig
						str_Warehouse_Exit = "NASHUA_ADJUSTMENT" ' Warehouse Enter
					Else
						str_Note = "B3: "
					End if

					Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)

				Else

					Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)

				End If

			End If

		Else

			Call PrefAddTransfer(pid, str_Warehouse_Exit, str_Warehouse_Entry, part, qty, color, LengthFt, False, "StockEdit", str_Note, po, str_Referrer)

		End If

	End If

'rs.close
'set rs=nothing
'rs2.close
'set rs2=nothing
'rs3.close
'set rs3=nothing
'rs4.close
'set rs4=nothing
'rs5.close
'set rs5=nothing
'DBConnection.close
'set DBConnection=nothing

DbCloseAll

End Function

%>

	</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockbyrackedit.asp?id=<% response.write pid %>&aisle=<% response.write aisle %>&ticket=<% response.write ticket%>&pobundleSEARCH=<% response.write pobundleSEARCH%>&poSEARCH=<% response.write poSEARCH%>&bundleSEARCH=<% response.write bundleSEARCH%>&part=<% response.write part%>" target="_self">Edit Stock</a>
    </div>

<form id="conf" title="Edit Stock" class="panel" name="conf" action="stockbyrackedit.asp" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

        <BR>

        <input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>">
		<input type="hidden" name='pobundleSearch' id='pobundleSearch' value="<%response.write pobundleSEARCH %>">
		<input type="hidden" name='bundleSearch' id='bundleSearch' value="<%response.write bundleSEARCH %>">
		<input type="hidden" name='poSearch' id='poSearch' value="<%response.write poSEARCH %>">
        <input type="text" name='part' id='part' value="<%response.write part %>">
		<input type="hidden" name='aisle' id='aisle' value="<%response.write aisle %>">
        <input type="hidden" name='id' id='id' value="<%response.write pid %>">
         <a class="whiteButton" href="javascript:conf.submit()" target="_self">Back to Stock</a>
          <a class="whiteButton" href="index.html#_Inv" target = "_self">Home</a>   
            </form>

</body>
</html>

