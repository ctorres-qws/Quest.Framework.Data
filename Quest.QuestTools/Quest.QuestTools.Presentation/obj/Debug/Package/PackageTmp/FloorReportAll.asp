                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!-- Collects information from FindStock.asp -->
		<!--Tool to show Employee, time, and Activity for Job/Floor/Tag
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>All Floor Info</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	
	</head>
	

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="FloorReportFind.asp" target="_self">Find Again</a>
        </div>
   


            
<ul id="screen1" title="Stock by JOB FLOOR TAG" selected="true">


<li class='group'>Windows</li>
<li><table border='1' class='sortable'><tr><th colspan = "5">Divided into Floor Groupings</th></tr>


<%
JOB =  Request.Querystring("JOB")
Floor = 0
Do while Floor < 100
FloorCountTotal = 0
HoistCountTotal = 0
Floor= Floor + 1

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT JOB, FLOOR, TAG, X, Y, Style FROM " & JOB & " WHERE FLOOR LIKE '" & FLOOR & "' OR FLOOR LIKE '" & FLOOR & "[!0-9]%'  ORDER by Floor"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
	
	
	if not rs2.eof then
	RS2.movefirst
	FloorName = UCASE(RS2("FLOOR"))
	response.write "<tr>"
	do while not rs2.eof
	FloorNameOld = FloorName
	FloorName = UCASE(RS2("FLOOR"))
	
	if FloorNameOld = FloorName then
		FloorCount = FloorCount + 1
	else
		
		response.write "<td><B>" & FloorNameOld & "</B>"
		response.write " " & FloorCount & "</td>"
		FloorCountTotal  = FloorCountTotal + FloorCount
			if FloorNameOld = Floor & "H" then
				HoistCountTotal = FloorCount
			end if
		FloorCount = 1
	
	end if 
	
	rs2.movenext
	loop
	
	
		
		response.write "<td><B>" & FloorName & "</B>"
		response.write " " & FloorCount & "</td>"
		FloorCountTotal  = FloorCountTotal + FloorCount
		if FloorNameOld = Floor & "H" then
		HoistCountTotal = FloorCount
		end if
		FloorCount = 0
		response.write "</tr>"
		response.write " Floor " & Floor & " Total Windows = " & FloorCountTotal
			if HoistCountTotal > 0 then
				response.write " Including Hoist " & Floor & " Total Windows = " & HoistCountTotal & "<br>"
			else
				response.write "<br>"
			end if
	end if
	rs2.close
	set rs2 = nothing
	
	
Loop
 
DBConnection.close
set DBConnection = nothing           
%>
</body>
</html>
