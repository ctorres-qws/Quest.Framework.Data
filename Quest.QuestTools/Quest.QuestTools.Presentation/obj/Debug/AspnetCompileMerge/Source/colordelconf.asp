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

cid = REQUEST.QueryString("ID")

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
                <a class="button leftButton" type="cancel" href="colordel.asp?id=<% response.write cid %>" target="_self">Delete Color</a>
    </div>

<%
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
cid = request.querystring("cid")
deletetype = request.querystring("del")
if deletetype = "on" then
	DEL = "1"
else
	DEL = "0"
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
strSQL = "SELECT * FROM Y_COLOR ORDER BY PROJECT ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

If DEL = "1" Then
	'Set Color Delete Statement
	StrSQL = FixSQLCheck("Delete * FROM Y_COLOR WHERE ID = " & Cid, isSQLServer)
	'Get a Record Set
	Set RS = DBConnection.Execute(strSQL)
Else
'Set Color Update Statement
	StrSQL = FixSQLCheck("UPDATE Y_COLOR SET ACTIVE=  FALSE WHERE ID = " & Cid, isSQLServer)
	'Get a Record Set
	Set RS = DBConnection.Execute(strSQL)
End If

'DBConnection.close
'set DBConnection=nothing	 

DbCloseAll

End Function

%>
		<form id="conf" title="Delete" class="panel" name="conf" action="colordel.asp" method="GET" target="_self" selected="true" >   

<% if DEL = "1" then %>
		
   <h2>Color: Deleted from Table</h2>
   <h2>Color Removed from Database</h2>
<% else %>   
		<h2>Color: Set to Inactive</h2>
		<h2>Color Removed from Choice list</h2>
<%end if%>

        <BR>

		<a class="whiteButton" href="javascript:conf.submit()">Back to Delete Colors</a>

            </form>
</body>
</html>

