<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodeqc.asp"-->

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

'Create a Query
    SQL = "Select * FROM X_BARCODE"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
	
'Create a Query
    SQL3 = "DELETE * FROM X_BARCODETEMPETV"
'Get a Record Set
    Set RS3 = DBConnection.Execute(SQL3)	
	
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODETEMPETV"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection



Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_EMPLOYEES"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODESW"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection

Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "SELECT * From X_BARCODESD"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection

Set rs8 = Server.CreateObject("adodb.recordset")
strSQL8 = "SELECT * From X_BARCODEOV"
rs8.Cursortype = 2
rs8.Locktype = 3
rs8.Open strSQL8, DBConnection
	

JOB = REQUEST.QueryString("JOB")
FL = REQUEST.QueryString("FLOOR")

totalg = 0
totala = 0
totalc = 0
TOTALZ = 0
totalsu = 0


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
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Assembly by Employee" selected="true">
    
    <li class="group">Today's Stats</li>
    
    <%

DO WHILE NOT RS5.EOF



RS.FILTER = "ID > 0"
rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
'rs.filter = "MONTH = " & cMonth & " AND YEAR = " & cYear
TOTALG = 0
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("EMPLOYEE") = RS5("NUMBER") AND RS("DEPT") = "ASSEMBLY" then
totalG = totalG + 1
response.write "<img src='greensquare.gif'>"
end if
rs.movenext
loop

IF TOTALG = 0 THEN 
ELSE
RESPONSE.WRITE "<li>"
%><a href="ETVadet.asp?employee=<% RESPONSE.WRITE RS5("NUMBER") %>&first=<% response.write rs5("first") %>&last=<% response.write rs5("last") %>" target="_SELF"><% RESPONSE.WRITE RS5("NUMBER") & " " & RS5("LAST") & "," & RS5("FIRST") & " : " & totaLG %></a></li><%
END IF 

RS5.MOVENEXT
LOOP

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
'rs3.close
'set rs3=nothing
DBConnection.close
set DBConnection=nothing
%>

