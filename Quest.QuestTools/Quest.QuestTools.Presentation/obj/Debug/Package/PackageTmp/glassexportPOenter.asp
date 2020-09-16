<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--
Designed Jan 2016 by Michael Bernholtz for Gurveen Sidhu and Sasha Bersovski to export to Optima based on PO
-->
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
                <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self" >Glass Tools</a>
    </div>

              <form id="edit" title="Select Service by PO" class="panel" name="edit" action="glassexportpo.asp" method="GET" target="_self" selected="true" >
        <h2>Search Goreway Stock by PO</h2>

<fieldset>
     <div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO'>
            </div>
            
</fieldset>

        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Optimize By PO</a><BR>

            </form> 

</body>
</html>

<%
'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection=nothing
%>

