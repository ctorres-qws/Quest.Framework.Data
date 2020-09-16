<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		 <!--Updated January 2014 to be both Table and Row form-->		
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


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'Torbram' ORDER BY AISLE, RACK, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

sortby = REQUEST.QueryString("sortby")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stocklevelsTorbram.asp" target="_self">Stock by Die</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Stock by Die" selected="true">

    

    <%
	part = request.QueryString("part")
	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
		InventoryType = rs2("InventoryType")
	end if
	Response.Write "<li class='group'><a href='stockbydieTorbramTable.asp?part=" & part & "' target='_self' >Stock (Row Form) - Switch to Table Form</a></li>"
	
	
	rs.movefirst
rs.filter = "PART = '" & part & "' AND WAREHOUSE='Torbram'"

response.write "<li><img src='/partpic/" & part & ".png'/> - " & Description & "</li>"
do while not rs.eof
po = rs("PO")
response.write "<li>" & rs.fields("PART") & " "
	if isnull(rs.fields("Colour")) then 
	response.write rs.fields("Project")
	else
	response.write rs.fields("Colour")
	end if
	
	Select Case InventoryType
	Case "Plastic"

		response.write " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " PO" & po & " / " & rs.fields("Bundle") & " " & rs.fields("Lmm") & " mm" &"</li>"
	
	Case "Sheet"
	
		response.write " Thickness " & Round(rs.fields("Thickness"),4) & " / " & rs.fields("Qty") & " SL" & " PO" & po &"</li>"

	Case Else 'Extrusion
		response.write " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " PO" & po & " / " & rs.fields("Bundle") & " " & rs.fields("Lmm") & " mm" &"</li>"
End Select
	
if not rs.fields("Allocation") = "" then
response.write "<li>- Allocated to: " & rs.fields("Allocation") & "</li>"
end if
'response.write "<li><a href='stockdet.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " " & rs.fields("Colour") & " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & "</a></li>"
'if color is missing put project
rs.movenext
loop

RESPONSE.WRITE "</UL>"

'rs6.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
'Do while not rs6.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not

'IF rs6("DEPT") = "GLASSLINE" then
'totalsu = totalsu + 1
'end if
 
 ' rs6.movenext
'loop
wcount=0
JFCHECKID=0
%>

</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs2=nothing


DBConnection.close
set DBConnection=nothing
%>

