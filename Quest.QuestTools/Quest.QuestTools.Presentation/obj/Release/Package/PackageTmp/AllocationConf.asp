<!--#include file="dbpath.asp"-->
                       <!-- Updated May 9th to include Length in Feet, Michael Bernhotlz -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		  <!--June 2016 - Shaun Levy has requested a list Y_INV_ADJ to record all adjustments made-->
		 <!-- This form is designed to REMOVE/ADD items from Inventory as an Adjustment and store a version of the Adjustment seperatly for accounting purposes.-->
		 <!-- This Form Affects Y_INV, Y_INVLOG, and Y_INV_ADJ -->
		 
		 
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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>

<%

ticket = Request.Querystring("ticket")
part = REQUEST.QueryString("part")
pid = request.querystring("id")
aisle = request.querystring("aisle")
FloorNote2 = request.querystring("FloorNote")

poSEARCH = request.querystring("poSEARCH")
bundleSEARCH = request.querystring("bundleSEARCH")
pobundleSEARCH = request.querystring("pobundleSEARCH")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & pid

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_INVLOG"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

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
WarehouseMove = Request.Querystring("WarehouseMove")
UpdateSuccess = False

allocation = REQUEST.QueryString("allocation")
length = REQUEST.QueryString("length")


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

color = REQUEST.QueryString("color")
aisle = REQUEST.QueryString("aisle")
if aisle = "I" then
	aisle = "i"
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


' Marker to Create DateOut when Warehouse Becomes Production for the first time
'July 15th, 2014 - Michael Bernholtz at Request of Shaun Levy
DateOut = ""
if UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION" then
	if UCASE(rs.Fields("warehouse")) <> "WINDOW PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "COM PRODUCTION" then 
		DateOut = currentDate
	end if 
end if

if (UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION") and (UCASE(warehouseMove) = "WINDOW PRODUCTION" or UCASE(warehouseMove) = "COM PRODUCTION") then
UpdateSuccess = FALSE
ErrorReason = "PROD"
'CODE added May 2016 to prevent Production items from being transfered within production.
'This type of transfer will alter daily production values and has to be cleared
else 	
	if QTY- QTYMOVE >0 then
	UpdateSuccess = TRUE

		rs2.AddNew
			rs2.Fields("Part") = rs.Fields("Part")
			rs2.Fields("colour") = rs.Fields("colour")
			rs2.Fields("qty") = rs.Fields("qty")
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
			rs2.Fields("exbundle") = rs.Fields("exbundle")
			rs2.Fields("thickness") = rs.Fields("thickness")
			rs2.Fields("transaction") = "original"
			rs2.Fields("day") = cday
			rs2.Fields("month") = cmonth
			rs2.Fields("year") = cyear
			rs2.Fields("week") = weeknumber
			rs2.Fields("time") = cctime
			rs2.Fields("ModifyDate") = currentDate
			rs2.fields("Itemid") = pid
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
			rs2.Fields("qty") = qty - qtyMove
			rs2.Fields("linch") = linch
			rs2.Fields("lmm") = lmm
			rs2.fields("lft") = lft
			rs2.Fields("warehouse") = warehouse
			rs2.Fields("aisle") = aisle
			rs2.Fields("rack") = rack
			rs2.Fields("shelf") = shelf
			rs2.Fields("po") = po
			rs2.Fields("colorpo") = colorpo
			rs2.Fields("bundle") = bundle
			rs2.Fields("exbundle") = exbundle
			rs2.Fields("thickness") = thickness
			rs2.Fields("Allocation") = allocation
			rs2.Fields("transaction") = "edit"
			rs2.Fields("day") = cday
			rs2.Fields("month") = cmonth
			rs2.Fields("year") = cyear
			rs2.Fields("week") = weeknumber
			rs2.Fields("time") = cctime
			rs2.Fields("ModifyDate") = currentDate
			rs2.fields("Note") = FloorNote
			if isDate("ExpectedDate") then
				rs2.Fields("ExpectedDate") = expdate
			end if
			if isDate(DateOut) then
				rs2.Fields("DateOut") = DateOut
			end if
			rs2.update

			rs.Fields("Part") = part
			rs.Fields("colour") = color
			rs.Fields("qty") = qty - QtyMove
			rs.Fields("linch") = linch
			rs.Fields("lmm") = lmm
			rs.fields("lft") = lft
			rs.Fields("warehouse") = warehouse
			rs.Fields("aisle") = aisle
			rs.Fields("rack") = rack
			rs.Fields("shelf") = shelf
			rs.Fields("po") = po
			rs.Fields("colorpo") = colorpo
			rs.Fields("bundle") = bundle
			rs.Fields("exbundle") = exbundle
			rs.Fields("thickness") = thickness
			rs.Fields("Allocation") = allocation
			rs.Fields("ModifyDate") = currentDate
			rs.fields("Note") = FloorNote
			if isDate(expdate) then
				rs.Fields("ExpectedDate") = expdate
			end if
			if isDate(DateOut) then
				rs.Fields("DateOut") = DateOut
			end if
			rs.update
			
		' -----------------------------------ADD NEW RECORD-------------------------------------------	
			
		
	rs.AddNew
		rs.Fields("Part") = part
		rs.Fields("colour") = color
		rs.Fields("qty") = qtyMove
		rs.Fields("linch") = linch
		rs.Fields("lmm") = lmm
		rs.Fields("lft") = lft
		rs.Fields("warehouse") = warehouseMOVE
		rs.Fields("PO") = PO
		rs.Fields("ColorPO") = colorPO
		rs.Fields("bundle") = BundleMove
		rs.Fields("aisle") = aisle
		rs.Fields("rack") = rack
		rs.Fields("shelf") = shelf
		rs.Fields("DateIn") = currentDate
		rs.Fields("bundle") = bundleMove
		rs.Fields("exbundle") = exbundle
		rs.Fields("thickness") = thickness
		rs.Fields("Allocation") = allocation
		rs.Fields("ModifyDate") = currentDate
		rs.fields("Note") = FloorNote
		
		if UCASE(warehouseMOVE) = "WINDOW PRODUCTION" or UCASE(warehouseMOVE) = "COM PRODUCTION" then
			rs.Fields("DateOut") = currentDate
		end if
		
		if isDate(expdate) then
			rs.Fields("ExpectedDate") = expdate
		end if
		rs.update

	rs2.AddNew
		rs2.Fields("Part") = part
		rs2.Fields("colour") = color
		rs2.Fields("qty") = qtyMove
		rs2.Fields("linch") = linch
		rs2.Fields("lmm") = lmm
		rs2.Fields("lft") = lft
		rs2.Fields("aisle") = aisle
		rs2.Fields("rack") = rack
		rs2.Fields("shelf") = shelf
		rs2.Fields("warehouse") = warehouseMove
		rs2.Fields("PO") = po
		rs2.Fields("colorPO") = colorpo
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
		rs2.Fields("Allocation") = allocation
		rs2.Fields("itemid") = pid
		rs2.fields("Note") = FloorNote
		
		if UCASE(warehouseMOVE) = "WINDOW PRODUCTION" or UCASE(warehouseMOVE) = "COM PRODUCTION" then
			rs2.Fields("DateOut") = currentDate
		end if
		
		
		if isDate(expdate) then
			rs2.Fields("ExpectedDate") = expdate
		end if

		rs2.update	

	UpdateSuccess = True	
	end if	
end if
rs.close
set rs=nothing
rs2.close
set rs2=nothing
DBConnection.close
set DBConnection=nothing
%>

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockbyrackedit.asp?id=<% response.write pid %>&aisle=<% response.write aisle %>&ticket=<% response.write ticket%>&pobundleSEARCH=<% response.write pobundleSEARCH%>&poSEARCH=<% response.write poSEARCH%>&bundleSEARCH=<% response.write bundleSEARCH%>" target="_self">Edit Stock</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="stockbyrackedit.asp?id=<% response.write pid %>&aisle=<% response.write aisle %>&ticket=<% response.write ticket%>&pobundleSEARCH=<% response.write pobundleSEARCH%>&poSEARCH=<% response.write poSEARCH%>&bundleSEARCH=<% response.write bundleSEARCH%>" method="GET" target="_self" selected="true" >              

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
        <input type="text" name='part' id='part' value="<%response.write part %>">
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
	 <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
	
	
	<%
	end if
%>

	
            </form>

            
    
</body>
</html>



