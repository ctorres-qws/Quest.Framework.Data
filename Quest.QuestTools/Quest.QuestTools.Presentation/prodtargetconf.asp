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
%>

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="prodtarget.asp" target="_self">Prod Target</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="index.html" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%

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

DAYG = REQUEST.QueryString("DAYG")
NIGHTG = REQUEST.QueryString("NIGHTG")
DAYA = REQUEST.QueryString("DAYA")
NIGHTA = REQUEST.QueryString("NIGHTA")
DAYIG = REQUEST.QueryString("DAYIG")
NIGHTIG = REQUEST.QueryString("NIGHTIG")

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
strSQL = "SELECT * FROM X_PRODTARGETS"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.movefirst
DO WHILE NOT RS.EOF
IF RS("Type") = "DAYG" then
rs.fields("Target") = DAYG
rs.update
end if 

IF RS("Type") = "DAYA" then
rs.fields("Target") = DAYA
rs.update
end if

IF RS("Type") = "DAYIG" then
rs.fields("Target") = DAYIG
rs.update
end if

IF RS("Type") = "NIGHTG" then
rs.fields("Target") = NIGHTG
rs.update
end if 

IF RS("Type") = "NIGHTA" then
rs.fields("Target") = NIGHTA
rs.update
end if

IF RS("Type") = "NIGHTIG" then
rs.fields("Target") = NIGHTIG
rs.update
end if
rs.movenext
loop

DbCloseAll

End Function

%>


        <BR>
       
        
        <input type="text" name='part' id='part' value="<%response.write "SYSTEM UPDATED" %>">
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
            
            </form>

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set conntemp=nothing
%>

