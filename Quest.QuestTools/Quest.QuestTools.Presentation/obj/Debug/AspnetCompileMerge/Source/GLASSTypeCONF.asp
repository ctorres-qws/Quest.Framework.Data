<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- On Submit of Manage Glass - page: glasstype.asp -->

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

GLASSTYPE = REQUEST.QueryString("GLASSTYPE")
DESCRIPTION = REQUEST.QueryString("DESCRIPTION")
SHOPCODE = REQUEST.QueryString("ShopCode")
STATUS = REQUEST.QueryString("Status")
JOB = REQUEST.QueryString("Job")

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

	Set rs5 = Server.CreateObject("adodb.recordset")
	strSQ5L = "Select * FROM XQSU_GlassTypes"
	rs5.Cursortype = GetDBCursorTypeInsert
	rs5.Locktype = GetDBLockTypeInsert
	rs5.Open strSQ5L, DBConnection

	rs5.AddNew
	rs5.Fields("Type") = GLASSTYPE
	rs5.Fields("Description") = DESCRIPTION
	rs5.Fields("Shopcode") = SHOPCODE
	rs5.Fields("Status") = STATUS
	rs5.Fields("Job") = JOB

	If GetID(isSQLServer,1) <> "" Then rs5.Fields("ID") = GetID(isSQLServer,1)
	rs5.update
	Call StoreID1(isSQLServer, rs5.Fields("ID"))

	DbCloseAll

End Function

	if STATUS = "" then
		STATUS = "Normal"
	end if
	
%>
	</head>
<body onload="startTime()" >

	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<a id="backButton" class="button" href="#"></a>
		<a class="button leftButton" type="cancel" href="glasstypes.asp" target="_self">Manage Glass</a>
		<a class="button" href="#searchForm" id="clock"></a>
	</div>

<ul id="Report" title="Added" selected="true">
	<li><% response.write "Type: " & GLASSTYPE %></li>
	<li><% response.write "Description: " & DESCRIPTION %></li>
	<li><% response.write "Shop Code: " & SHOPCODE %></li>
	<li><% response.write "Status: " & STATUS %></li>
	<li><% response.write "Job: " & JOB %></li>
</ul>

<%
//rs5.close
//set rs5=nothing

//DBConnection.close
//set DBConnection=nothing
%>

</body>
</html>

