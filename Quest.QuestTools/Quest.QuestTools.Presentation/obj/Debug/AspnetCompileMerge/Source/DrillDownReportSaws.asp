<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Glass Reporting - Shows the Glass cut based on JOB - Total, This Month, Today -->
<!-- Forel and Willian, Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Drill Down Stats</title>
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

<!--#include file="todayandyesterday.asp"-->

</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
		<% 
			Ticket = Request.QueryString("Ticket") 
			If Ticket = "BarcoderTV" then
			BackButton = "BarcoderTV.asp"
			Else
			BackButton = "index.html#_Report"
			End if
			

			
			
			
			
			
		%>
                <a class="button leftButton" type="cancel" href="<%Response.Write BackButton %>" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Saw Drill Down Stats" selected="true">

		<li class="group">Extrusion Drill Down</li>
<%

Set rs = Server.CreateObject("adodb.recordset")
Today = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
strSQL = FixSQL("SELECT * From PROECOHOR where StartDate = # "& Today & "# ORDER BY JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Response.write "<li class='group'>Today (Ecowall Horizontal)</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	do while not rs.eof
	
	response.write "<li>" & rs("JobNumber") & ": " & rs("CutStatus") & "%</li>"

	
	rs.movenext
	loop
end if
rs.close
set rs=nothing

Set rs = Server.CreateObject("adodb.recordset")
Today = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
strSQL = FixSQL("SELECT * From PROECOVERT where StartDate = # "& Today & "# ORDER BY JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Response.write "<li class='group'>Today (Ecowall Vertical)</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	do while not rs.eof
	
	response.write "<li>" & rs("JobNumber") & ": " & rs("CutStatus") & "%</li>"

	
	rs.movenext
	loop
end if
rs.close
set rs=nothing

Set rs = Server.CreateObject("adodb.recordset")
Today = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
strSQL = FixSQL("SELECT * From PROQHOR where StartDate = # "& Today & "# ORDER BY JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Response.write "<li class='group'>Today (Q4750 Horizontal)</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	do while not rs.eof
	
	response.write "<li>" & rs("JobNumber") & ": " & rs("CutStatus") & "%</li>"

	
	rs.movenext
	loop
end if
rs.close
set rs=nothing


Set rs = Server.CreateObject("adodb.recordset")
Today = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
strSQL = FixSQL("SELECT * From PROQVERT where StartDate = # "& Today & "# ORDER BY JobNumber ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Response.write "<li class='group'>Today (Q4750 Vertical)</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	do while not rs.eof
	
	response.write "<li>" & rs("JobNumber") & ": " & rs("CutStatus") & "%</li>"

	
	rs.movenext
	loop
end if
rs.close
set rs=nothing

%>

	</ul>
        
  
<% 

DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
