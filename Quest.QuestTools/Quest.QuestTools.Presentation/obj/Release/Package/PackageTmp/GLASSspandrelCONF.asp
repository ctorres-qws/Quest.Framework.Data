<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted July 31st, 2014 - by Michael Bernholtz at request of Jody Cash --> 
<!-- On Submit of Manage Glass - page: glassspandrel.asp -->
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

SCODE = REQUEST.QueryString("CODE")
DESCRIPTION = REQUEST.QueryString("DESCRIPTION")
JOB = REQUEST.QueryString("JOB")
NOTES = REQUEST.QueryString("NOTES")
ACTIVE = REQUEST.QueryString("ACTIVE")

If ACTIVE = "on" Then
	ACTIVE = TRUE
Else
	ACTIVE = FALSE
End If

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

	'StrSQL = FixSQL("INSERT INTO Y_COLOR_SPANDREL (CODE, DESCRIPTION, JOB, NOTES, ACTIVE ) VALUES ('" & SCODE & "', '" & DESCRIPTION & "', '" & JOB & "', '" & NOTES & "', " & ACTIVE & " )")
	'Set RS = DBConnection.Execute(strSQL)

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Y_COLOR_SPANDREL WHERE ID=-1"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection

	RS.AddNew
	RS.Fields("CODE") = SCODE
	RS.Fields("DESCRIPTION") = DESCRIPTION
	RS.Fields("JOB") = JOB
	RS.Fields("NOTES") = NOTES
	RS.Fields("ACTIVE") = ACTIVE

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	RS.Update

	Call StoreID1(isSQLServer, rs.Fields("ID"))

	'DBConnection.close
	'set DBConnection=nothing

	DbCloseAll
End Function

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="glassspandrel.asp" target="_self">Manage Spandrel</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="Report" title="Added" selected="true">

    <li><% response.write "Code: " & SCODE %></li>
	<li><% response.write "Description: " & DESCRIPTION %></li>
	<li><% response.write "Job: " & Job %></li>
	<li><% response.write "Notes: " & NOTES %></li>
	<li><% response.write "ACTIVE: " & Active %></li>

</ul>

</body>
</html>

