<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Optimization Log Information presented in Report form-->
<!-- Reuqested by Victor and designed by Michael Bernholtz, August 2014 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optimization Log Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

    <%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM OptimizeLog ORDER BY GlassCutDate DESC, OpDate DESC, Glass"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
       <a class="button leftButton" type="cancel" href="index.html#_GlassP" target="_self">Glass Prod</a>
        </div>
  
        <ul id="Profiles" title="Optimization Log - Report" selected="true">
<% 


	
	

response.write "<li class='group'>Lites to be Cut Analysis</li>"
response.write "<li><table border='1' class='sortable' ><tr><th>Date</th><th>Glass</th><th># of Lites to be Cut</th></tr>"
rs.filter = ""
rs.filter = "GlassCutDate = NULL "
if rs.eof then
	response.write "<tr><td colspan = 3> Empty Table</td></tr>"
else

	GlassType = RS("Glass")
	GlassType2 = RS("Glass")
	GlassDate = RS("OpDate")
	GlassDate2 = RS("OpDate")
	TypeNum = 0
	do while not rs.eof

	GlassType = RS("Glass")
	GlassDate = RS("OpDate")
	if GlassType = GlassType2 and GlassDate = GlassDate2 then
		TypeNum = TypeNum + rs("Lites") + 0
	else
		response.write "<tr>"
		response.write "<td>" & GlassDate2 & "</td>"
		response.write "<td>" & GlassType2 & "</td>"
		response.write "<td>" & TypeNum & "</td>"
		response.write "</tr>" 

		TypeNum =  rs("Lites") + 0
	
	end if 

	GlassType2 = GlassType
	GlassDate2 = GlassDate
		rs.movenext
	loop

	response.write "<tr>"
	response.write "<td>" & GlassDate2 & "</td>"
	response.write "<td>" & GlassType2 & "</td>"
	response.write "<td>" & TypeNum & "</td>"
	response.write "</tr>"
end if
response.write "</table></li>"

response.write "<li class='group'>Lites Cut Analysis</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Date</th><th>Shift</th><th># of Lites Cut</th></tr>"
rs.filter = "GlassCutDate <> NULL "
if rs.eof then
	response.write "<tr><td colspan = 3> Empty Table</td></tr>"
else
	GlassDate = RS("GlassCutDate")
	GlassDate2 = RS("GlassCutDate")
	ShiftNight = 0
	ShiftDay =0

	do while not rs.eof

	GlassDate = RS("GlassCutDate")
	if GlassDate = GlassDate2 then
		if rs("SHIFT") = "NightShift" then
			ShiftNight = ShiftNight + rs("Lites") +0
		end if
		if rs("SHIFT") = "DayShift" then
			ShiftDay = ShiftDay + rs("Lites") + 0
		end if
	
	else
		response.write "<tr>"
		response.write "<td>" & GlassDate2 & "</td>"
		response.write "<td>Day Shift</td>"
		response.write "<td>" & ShiftDay & "</td>"
		response.write "</tr>" 
		response.write "<tr>"
		response.write "<td>" & GlassDate2 & "</td>"
		response.write "<td>Night Shift</td>"
		response.write "<td>" & ShiftNight & "</td>"
		response.write "</tr>" 
		if rs("SHIFT") = "NightShift" then
			ShiftNight = rs("Lites")
		end if
			if rs("SHIFT") = "DayShift" then
			ShiftDay =  rs("Lites")
		end if
	
	end if 
	GlassDate2 = GlassDate
		rs.movenext
	loop
	response.write "<tr>"
	response.write "<td>" & GlassDate2 & "</td>"
	response.write "<td>Day Shift</td>"
	response.write "<td>" & ShiftDay & "</td>"
	response.write "</tr>" 
	response.write "<tr>"
	response.write "<td>" & GlassDate2 & "</td>"
	response.write "<td>Night Shift</td>"
	response.write "<td>" & ShiftNight & "</td>"
	response.write "</tr>" 
	
	response.write "</table></li>"
	
end if

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>        
            
         
               
</body>
</html>
