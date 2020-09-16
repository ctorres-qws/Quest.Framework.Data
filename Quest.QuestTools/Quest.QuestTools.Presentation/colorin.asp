<!--#include file="dbpath.asp"-->                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Updated January 2015, Michael Bernholtz, to split Job and Side rather than a single field. This will help with Database Consistency -->
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
	
JOB = TRIM(REQUEST.QueryString("JOB"))
SIDE = TRIM(REQUEST.QueryString("SIDE"))
Project = TRIM(JOB & " " & SIDE)
CODE = TRIM(REQUEST.QueryString("CODE"))
COMPANY = TRIM(REQUEST.QueryString("COMPANY"))
DES = TRIM(REQUEST.QueryString("DESCRIPTION"))

PAINTCAT = TRIM(REQUEST.QueryString("PAINTCAT"))
ACTIVE = REQUEST.QueryString("ACTIVE")
if ACTIVE = "on" then
	ACTIVE = TRUE
else
	ACTIVE = FALSE
end if
EXTRUSION = REQUEST.QueryString("EXTRUSION")
if EXTRUSION = "on" then
	EXTRUSION = TRUE
else
	EXTRUSION = FALSE
end if
SHEET = REQUEST.QueryString("SHEET")
if SHEET = "on" then
	SHEET = TRUE
else
	SHEET = FALSE
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
strSQL = "SELECT * FROM Y_COLOR"
rs.Cursortype = GetDBCursorTypeInsert
rs.Locktype = GetDBLockTypeInsert
rs.Open strSQL, DBConnection

rs.movelast
if  rs.Fields("PROJECT") = PROJECT AND rs.Fields("JOB") = JOB AND rs.Fields("SIDE") = SIDE then
Fail =1
else
rs.AddNew
	rs.Fields("PROJECT") = PROJECT
	rs.Fields("JOB") = JOB
	rs.Fields("CODE") = CODE
	rs.Fields("COMPANY") = COMPANY
	rs.Fields("DESC") = DES
	rs.Fields("SIDE") = SIDE
	rs.Fields("PRICECAT") = PAINTCAT
	rs.Fields("ACTIVE") = ACTIVE
	rs.Fields("EXTRUSION") = EXTRUSION
	rs.Fields("SHEET") = SHEET

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	rs.update

	Call StoreID1(isSQLServer, rs.Fields("ID"))

end if

DbCloseAll

End Function

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="COLORadd.asp#_enter" target="_self">Add Color</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

    
<ul id="Report" title="Added" selected="true">
	<%if Fail =1 then 
	 response.write "<li> Existing Colour Not Added Again:</li>"
	else
	%>
    <li> New Colour Added:<%response.write SessionEntry%></li>
	<li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Side: " & SIDE %></li>
	<li><% response.write "Paint Code: " & CODE & " at " & COMPANY %></li>
	<li><% response.write "Paint Location: " & DES %></li>
    <li><% response.write "Price Catagory: " & PAINTCAT %></li>
	<li><% response.write "ACTIVE: " & ACTIVE %></li>
	<%
	if EXTRUSION = TRUE then
	response.write "<li>Colour For: Extrusion</li>"
	end if
	%>
	<%
	if SHEET = TRUE then
	response.write "<li>Colour For: Sheet</li>"
	end if
	%>
	
	<%
	end if
	%>
	
	</li>
  
  <a class="whiteButton" href="coloradd.asp" target ="_Self">Back to Add Colors</a>
</ul>



<% 
On Error Resume Next
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>
