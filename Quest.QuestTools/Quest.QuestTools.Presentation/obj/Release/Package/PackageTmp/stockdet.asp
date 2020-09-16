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


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

part = request.QueryString("part")

id = REQUEST.QueryString("ID")


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
                <a class="button leftButton" type="cancel" href="stockbydie.asp?part=<% response.write part %>" target="_self">Stock by Die</a>
    </div>

<ul id="screen1" title="Detailed View" selected="true">
    
    <li class="group">Location</li>

    <%
	rs.filter = "ID = '" & id & "'"
'response.write "<li>IMAGE LINKN !!!!!!!: " & rs.fields("AISLE") & "</li>"
response.write "<li><img src='/partpic/" & rs.fields("PART") & ".png'/></li>"
response.write "<li>AISLE: " & rs.fields("AISLE") & "</li>"
response.write "<li>RACK: " & rs.fields("RACK") & "</li>"
response.write "<li>SHELF: " & rs.fields("SHELF") & "</li>"


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

DBConnection.close
set DBConnection=nothing
%>

