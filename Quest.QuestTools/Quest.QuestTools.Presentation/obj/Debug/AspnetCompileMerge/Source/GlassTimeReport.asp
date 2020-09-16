<!--#include file="dbpath.asp"-->
                       
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM V_REPORT1 ORDER BY ID DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
rs.filter = "DAY = " & sDAY & " AND MONTH =" & SMonth & " AND YEAR = " & SYear 
DayID = rs("ID")
rs.filter = ""


%>

</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
				<a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

	
		<li class="group">Glass Line Statistics</li>

<%
rs.filter = "ID >= " & DayID
Response.write "<li> Click on a Header to Sort by that Column</li>"
Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
Response.write" <tr><th>Date</th><th>Forel</th><th>Willian</th><th>Total</th></tr>"

Do while not rs.eof

	GDay = rs("DAY")
	GMonth = rs("MONTH")
	GYear = rs("YEAR")
	Forel = rs("Forel") + 0
	Willian = rs("Willian") + 0
	Total = Forel + Willian + 0
	
	Response.write "<tr><td>" & GDay & "/" & GMonth & "/" & GYear & "</td><td> " & Forel & "</td><td> " & Willian & "</td><td> " & Total & "</td></tr>"
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

