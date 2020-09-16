<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="masteredittable.asp?part=<% response.write part %>" target="_self">Edit Master</a>
    </div>

<form id="conf" title="Edit Stock" class="panel" name="conf" action="masteredittable.asp#_screen1" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%
pid = request.querystring("id")

'Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL = "SELECT * From Y_INVLOG"
'rs2.Cursortype = 2
'rs2.Locktype = 3
'rs2.Open strSQL, DBConnection

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

part = request.querystring("part")
description = request.querystring("description")
supplierpart = REQUEST.QueryString("supplierpart")
kgm = REQUEST.QueryString("kgm")
if kgm = "" then
kgm = 0
end if
lbf = REQUEST.QueryString("lbf")
if lbf = "" then
lbf = 0
end if
MinLevel = request.querystring("MinLevel")
if MinLevel = "" then
	MinLevel = 250
	'Default Value
end if
paintcat = REQUEST.QueryString("paintcat")
inventorytype = REQUEST.QueryString("inventorytype")

if inventorytype = "Plastic" then
	Min16 = REQUEST.QueryString("Min-16")
	Min18 = REQUEST.QueryString("Min-18")
	Min20 = REQUEST.QueryString("Min-20")
	Min21 = REQUEST.QueryString("Min-21")
	Min22 = REQUEST.QueryString("Min-22")
	
	if Min16 = "" then
		Min16 = 250
		'Default Value
	end if
	if Min18 = "" then
		Min18 = 250
		'Default Value
	end if
	if Min20 = "" then
		Min20 = 250
		'Default Value
	end if
	if Min21 = "" then
		Min21 = 250
		'Default Value
	end if
	if Min22 = "" then
		Min22 = 250
		'Default Value
	end if
End if

HYDRO = REQUEST.QueryString("HYDRO")
canart = REQUEST.QueryString("canart")
keymark = REQUEST.QueryString("keymark")
extal = REQUEST.QueryString("extal")

if length > 300 then
linch = length / 25.4
lmm = length
end if

if length < 100 then
linch = length * 12
lmm = linch * 25.4
else
linch = length
lmm = linch * 25.4
end if

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
	strSQL = "SELECT * FROM Y_MASTER"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "ID = " & pid

	
	if rs.fields("Part") = part then
	else
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT * FROM Y_INV WHERE PART = '" & RS.fields("PART") & "'"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		
		DO while not rs2.EOF
			rs2.Fields("Part") = part
			rs2.update
		RS2.movenext
		loop
		
	end if
	
	
	
	rs.Fields("Part") = part
	rs.Fields("Description") = Description
	rs.Fields("supplierpart") = supplierpart
	rs.Fields("kgm") = kgm
	rs.Fields("lbf") = lbf
	rs.Fields("MinLevel") = MinLevel
	rs.Fields("paintcat") = paintcat
	rs.Fields("inventorytype") = inventorytype
	rs.Fields("HYDRO") = HYDRO
	rs.Fields("canart") = canart
	rs.Fields("keymark") = keymark
	rs.Fields("extal") = extal

	if inventoryType = "Plastic" then
	rs.Fields("Min-16") = Min16
	rs.Fields("Min-18") = Min18
	rs.Fields("Min-20") = Min20
	rs.Fields("Min-21") = Min21
	rs.Fields("Min-22") = Min22
	end if

	rs.update

	DbCloseAll
End Function

%>

        <BR>
       
        
        <input type="text" name='part' id='part' value="<%response.write part %>">
         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
            
            </form>

            
    
</body>
</html>

<% 

'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection=nothing
%>

