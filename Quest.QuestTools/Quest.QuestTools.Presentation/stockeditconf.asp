<!--#include file="dbpath.asp"-->
                       
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
part = request.querystring("part")

pid = request.querystring("id")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV "
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & pid

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_INVLOG"
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

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


colour = REQUEST.QueryString("colour")
qty = REQUEST.QueryString("qty")
length = REQUEST.QueryString("length")

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
if aisle = "in" or aisle = "In" or aisle = "IN" or aisle = "Inside" or aisle = "INSIDE"then
	aisle = "inside"
end if
if aisle = "out" or aisle = "Out" or aisle = "OUT" or aisle = "Outside" or aisle = "OUTSIDE"then
	aisle = "outside"
end if

currentDate = Date()

rack = REQUEST.QueryString("rack")
shelf = REQUEST.QueryString("shelf")
warehouse = REQUEST.QueryString("warehouse")
'Added po at Request of Ruslan - January 16, Michael Bernholtz
po = REQUEST.QueryString("po")
colorpo = REQUEST.QueryString("ColorPO")
' Added Bundle and External Bundle for Shaun, June-July 2014, Michael Bernholtz
bundle = REQUEST.QueryString("bundle")
exbundle = REQUEST.QueryString("exbundle")
allocation = REQUEST.QueryString("allocation")
thickness = REQUEST.QueryString("thickness")
FloorNote = ""
FloorNote = REQUEST.QueryString("FloorNote")
StatusNote = REQUEST.QueryString("StatusNote")
if thickness = "" then
	thickness = 0
end if
if isDate(rs.Fields("ExpectedDate")) then
		rs2.Fields("ExpectedDate") = rs.Fields("ExpectedDate")
end if


' Marker to Create DateOut when Warehouse Becomes Production for the first time
'July 15th, 2014 - Michael Bernholtz at Request of Shaun Levy
DateOut = ""
if UCASE(warehouse) = "WINDOW PRODUCTION" or UCASE(warehouse) = "COM PRODUCTION" then
	if UCASE(rs.Fields("warehouse")) <> "WINDOW PRODUCTION" AND UCASE(rs.Fields("warehouse")) <> "COM PRODUCTION" then 
		DateOut = currentDate
	end if 
end if

rs2.AddNew
	rs2.Fields("Part") = rs.Fields("Part")
	rs2.Fields("colour") = rs.Fields("colour")
	rs2.Fields("qty") = rs.Fields("qty")
	rs2.Fields("linch") = rs.Fields("linch")
	rs2.Fields("lmm") = rs.Fields("lmm")
	rs2.Fields("lft") = rs.Fields("lft")
	rs2.Fields("warehouse") = rs.Fields("warehouse")
	'Added po at Request of Ruslan - January 16, Michael Bernholtz
	rs2.Fields("po") = rs.Fields("po")
	rs2.Fields("colorpo") = rs.Fields("colorpo")
	rs2.Fields("bundle") = rs.Fields("bundle")
	rs2.Fields("exbundle") = rs.Fields("exbundle")
	rs2.Fields("allocation") = rs.Fields("allocation")
	rs2.Fields("thickness") = rs.Fields("thickness")
	rs2.Fields("aisle") = rs.Fields("aisle")
	rs2.Fields("rack") = rs.Fields("rack")
	rs2.Fields("shelf") = rs.Fields("shelf")
	rs2.Fields("ItemId") = rs.Fields("ID")
	rs2.Fields("PREF") = rs.Fields("PREF")
	rs2.Fields("Note") = rs.Fields("Note")
	rs2.Fields("Note 2") = rs.Fields("Note 2")
	rs2.Fields("transaction") = "original"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ModifyDate") = currentDate
	if isDate(rs.Fields("DateOut")) then
		rs2.Fields("DateOut") = rs.Fields("DateOut")
	end if
	rs2.update

rs2.AddNew
	rs2.Fields("Part") = part
	rs2.Fields("colour") = color
	rs2.Fields("qty") = qty
	rs2.Fields("linch") = linch
	rs2.Fields("lmm") = lmm
	rs2.Fields("lft") = lft
	rs2.Fields("warehouse") = warehouse
	rs2.Fields("po") = po
	rs2.Fields("colorpo") = colorpo
	rs2.Fields("bundle") = bundle
	rs2.Fields("exbundle") = exbundle
	rs2.Fields("allocation") = allocation
	rs2.Fields("thickness") = thickness
	rs2.Fields("aisle") = aisle
	rs2.Fields("rack") = rack
	rs2.Fields("shelf") = shelf
	rs2.Fields("transaction") = "edit"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ModifyDate") = currentDate
	rs2.Fields("ItemId") = pid
	rs2.Fields("Note") = FloorNote
	rs2.Fields("Note 2") = statusNote
	if isDate(DateOut) then
		rs2.Fields("DateOut") = DateOut
	end if
	rs2.update

	rs.Fields("Part") = part
	rs.Fields("colour") = color
	rs.Fields("qty") = qty
	rs.Fields("linch") = linch
	rs.Fields("lmm") = lmm
	rs.Fields("lft") = lft
	rs.Fields("warehouse") = warehouse
	rs.Fields("po") = po
	rs.Fields("colorpo") = colorpo
	rs.Fields("bundle") = bundle
	rs.Fields("exbundle") = exbundle
	rs.Fields("allocation") = allocation
	rs.fields("thickness") = thickness
	rs.Fields("aisle") = aisle
	rs.Fields("rack") = rack
	rs.Fields("shelf") = shelf
	rs.Fields("Note") = FloorNote
	rs.Fields("Note 2") = statusNote
	rs.Fields("ModifyDate") = currentDate
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
	
	
rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing
rs4.close
set rs4=nothing


DBConnection.close
set DBConnection=nothing
%>

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockedit.asp?part=<% response.write part %>" target="_self">Edit Stock</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="stockedit.asp#_screen1" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>



        <BR>
       <ul>
	   <li>Part: <%response.write part %></li>
	   <li>QTY: <%response.write qty %></li>
	   <li>Warehouse: <%response.write warehouse %></li>
	   <li>PO: <%response.write po %></li>
	   
	   </ul>
      
         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
            
            </form>

            
    
</body>
</html>


