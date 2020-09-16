<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- On Submit of Manage Glass - page: glassspacers.asp -->

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

SPACER = REQUEST.QueryString("SPACER")
OT = REQUEST.QueryString("OT")

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

	Set rs6 = Server.CreateObject("adodb.recordset")
	strSQL6 = "Select * FROM XQSU_OTSpacer WHERE ID=-1"
	rs6.Cursortype = GetDBCursorTypeInsert
	rs6.Locktype = GetDBLockTypeInsert
	rs6.Open strSQL6, DBConnection

	rs6.AddNew
	rs6.Fields("Spacer") = SPACER
	rs6.Fields("OT") = OT

	If GetID(isSQLServer,1) <> "" Then rs6.Fields("ID") = GetID(isSQLServer,1)
	rs6.update
	Call StoreID1(isSQLServer, rs6.Fields("ID"))

	DbCloseAll

End Function

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="glassspacers.asp" target="_self">Manage Spacers</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Spacer " & SPACER %></li>
	<li><% response.write "OT " & OT %></li>

</ul>

<%

rs6.close
set rs6=nothing

DBConnection.close
set DBConnection=nothing
%>

</body>
</html>


