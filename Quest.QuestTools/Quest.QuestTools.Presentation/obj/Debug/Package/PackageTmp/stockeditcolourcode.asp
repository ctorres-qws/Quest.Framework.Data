<!--#include file="dbpath.asp"-->

<!-- Duplicate form of Stockedit.asp changed to organize by Colour Code instead of by Colour-->   
<!-- Created August 1st, 2014, With permission by Jody Cash, by Michael Bernholtz-->  

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


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY Colour ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART, WAREHOUSE, COLOUR"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



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
                <a class="button leftButton" type="cancel" href="stock.asp#_colour" target="_self">Stock by Colour</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Stock by Die" selected="true">
    
    <li class="group">Stock</li>
    

    <%
	
	colour = request.QueryString("colour")
	rs.movefirst
rs.filter = "COLOUR = '" & colour & "' AND WAREHOUSE <> 'WINDOW PRODUCTION' AND WAREHOUSE <> 'COM PRODUCTION'"
do while not rs.eof
response.write "<li><a href='stockeditform.asp?id=" & rs.fields("ID") & "&part=" & rs.fields("PART") & "' target='_self'>" & rs.fields("PART") & " "
'if color is missing put project
if isnull(rs.fields("Colour")) then 
	response.write rs.fields("Project")
	else
	response.write rs.fields("Colour")
end if
'Added Lmm (length in mm at Request of Ruslan - January 16, Michael Bernholtz
response.write " PO " & rs.fields("po") & " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " " & rs.fields("Lmm") & " mm" & " IN: " & rs.fields("warehouse") & " Entered: " & rs.fields("DateIn") & "</a></li>"
if rs.fields("warehouse") = "GOREWAY" then
response.write "<li>- Aisle " & rs.fields("Aisle") & " Rack " & rs.fields("rack") & " Shelf " & rs.fields("shelf") & "</li>"
end if
rs.movenext
loop

RESPONSE.WRITE "</UL>"

%>

</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

