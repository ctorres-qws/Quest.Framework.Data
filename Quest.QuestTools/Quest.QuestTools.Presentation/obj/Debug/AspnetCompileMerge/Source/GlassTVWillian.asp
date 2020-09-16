<!--#include file="dbpath.asp"-->
                       
<!--Requested by Alex Stamenkovic -->	
<!-- Only shows Willian -->				   
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  <script src="sorttable.js"></script>
  
  
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

	sDay = trim(Request.Querystring("sDay"))
	sMonth = trim(Request.Querystring("sMonth"))
	sYear = trim(Request.Querystring("sYear"))
	
if sDay = "" or sMonth = "" or sYear = "" then

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
sDay = day(now)
sMonth = month(now)
sYear= year(now)
else

STAMPVAR = sMonth & "/" & sDay & "/" & sYear

end if

Call SetTestDate(sDay, sMonth, sYear)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEGA WHERE DAY = " & sDAY & " AND MONTH =" & SMonth & " AND YEAR = " & SYear &"  AND DEPT = 'WILLIAN' ORDER BY DATETIME DESC"

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

totalg = 0
totalgw = 0
totals = 0


Do while not rs.eof

	DATETIME = rs("DATETIME")
	GDay = rs("DAY")
	GMonth = rs("MONTH")
	GYear = rs("YEAR")

			IF UCASE(rs("DEPT")) = "WILLIAN" then
				totalg = totalg + 1
				totalgw = totalgw + 1
			end if

			IF UCASE(LEFT(rs("BARCODE"),2)) = "GT" then
				totals = totals + 1
			end if


rs.movenext
loop

%>

</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="glassTV.asp" target="_self">Glass</a>
				<a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Today's Production</li>
		<li><a href = "GlassTV.asp" target= "_self"> <% response.write "Total Insulated Glass: " & totalg %></a></li>
		<li><% response.write "Willian Insulated Glass: " & totalgw %></li>
		<li><% response.write "Service Coded Glass: " & totals %></li>
		
		<li class="group">Today's Scans</li>

<%
rs.filter = ""

	Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
	Response.write "<tr title = 'Click on a Header to Sort by that Column' ><th>Job</th><th>Floor</th><th>Tag</th><th>Opening #</th><th>Type</th><th>PO</th><th>PO Line #</th><th>Department</th><th>TimeStamp</th><th>Barcode</th></tr>"
	

Do while not rs.eof
	DATETIME = rs("DATETIME")
	
				response.write "<tr>"
				response.write "<td>" & rs("JOB") & "</td>"
				response.write "<td>" & rs("FLOOR") & "</td>"
				response.write "<td>" & rs("Tag") & "</td>"
				response.write "<td>" & rs("POSITION") & "</td>"
				response.write "<td>" & rs("Type") & "</td>"
				response.write "<td>" & rs("PO") & "</td>"
				response.write "<td>" & rs("POLINE") & "</td>"
				response.write "<td>" & rs("DEPT") & "</td>"
				response.write "<td>" & rs("DATETIME") & "</td>"
				response.write "<td>" & rs("BARCODE") & "</td>"
				response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

