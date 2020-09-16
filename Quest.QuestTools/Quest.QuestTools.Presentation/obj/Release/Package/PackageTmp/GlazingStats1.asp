 
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Glazing Statistics for Scanning based on Job and Floor-->
<!-- THis page shows Glazing information by Job as a whole, but when you put in Job and Floor it will run through each window to describe Complete / Partial-->
<!-- Michael Bernholtz, April 2016, at Request of Shaun Levy and Jody Cash -->
	<!--#include file="dbpath.asp"-->	 
		
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Cut File Progress</title>
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
JobFloor = False
AllWindows = 0
AllScans = 0
Job = request.querystring("Job")
Floor = request.querystring("Floor")
  
	if Job = "" then	
		Job = "ALL"
	else
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "Select * FROM " & JOB & " order by TAG ASC"
	rs2.Cursortype = GetDBCursorType
	rs2.Locktype = GetDBLockType
	rs2.Open strSQL2, DBConnection
	end if
		if Floor = "" then	
		Floor = "ALL"
	end if
  if JOB = "ALL" then
	 strSQL = "Select * FROM X_GLAZING order by BARCODE ASC"
  else
		if Floor = "ALL" then
			strSQL = "Select * FROM X_GLAZING Where JOB = '" & JOB & "' order by BARCODE ASC"
		else
			strSQL = "Select * FROM X_GLAZING Where JOB = '" & JOB & "' and FLOOR = '" & Floor & "' order by BARCODE ASC"
			rs2.filter = "FLOOR = '" & Floor & "'" 
			JobFloor = True
		end if
		
		Do while not rs2.eof
		AllWindows = AllWindows + 1
		rs2.movenext
		loop
  end if 

Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


ScanCompleteWindow = 0
ScanWindow = 0

Do while not rs.eof
OldBarcode = Barcode
Barcode = rs("Barcode")
	AllScans = AllScans + 1	
	if OldBarcode = Barcode then
		if rs("FirstComplete") = "TRUE" then
			ScanCompleteWindow = ScanCompleteWindow + 1
			ScanWindow = ScanWindow - 1
		end if
	else
		if rs("FirstComplete") = "TRUE" then
			ScanCompleteWindow = ScanCompleteWindow + 1
		else
			ScanWindow = ScanWindow + 1
		end if
	end if
rs.movenext
loop



%>  


</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>
<form id="screen1" title="Glaze Stats" class="panel" name="Glaze" action="GlazingStats1.asp" method="GET" selected="true">
<fieldset>
   <div class="row">   
            <label>Job </label>
            <input type="text" name='Job' id='Job' value = "<%response.write Job%>" >
		</div>
	<div class="row">     
            <label>Floor </label>
            <input type="text" name='Floor' id='Floor' value = "<%response.write Floor%>">
		</div>
		<a class="whiteButton" onClick=" Glaze.submit()">View Glazing Statistics</a><BR>
</fieldset>

	<ul id="screen1" title="Glazing Today" selected="true">

		<li class="group">All Glazing for <%response.write JOB & " " & Floor%> </li>
			<li><% response.write "GLAZING Complete: " & ScanCompleteWindow %></li>
			<li><% response.write "GLAZING Partial: " & ScanWindow %></li>
			<li><% response.write "Out of Windows: " & AllWindows %></li>
			<li><% response.write "Total Scans(Partial and Complete): " & AllScans %></li>

	</ul>
	
	<%
	If JobFloor = TRUE then
	
	%>
	
		<ul id="screen2" title="Windows" selected="true">

		<li class="group">All Windows for <%response.write JOB & " " & Floor%> </li>
		
		<%
		rs2.filter = "FLOOR = '" & Floor & "'" 
		if not rs2.eof then
		rs2.movefirst 
		else
		%>
		<li>Empty File</li>
		<%
		end if
		Do while not rs2.eof 
		JobBarcode = RS2("JOB") & RS2("FLOOR") & RS2("TAG")
		rs.filter = "BARCODE = '" & JobBarcode & "'"
		if rs.eof then
		WindowStatus = "Not Scanned"
		else
		WindowStatus = "Partial"
		Do while not rs.eof
			if rs("FirstComplete") = "TRUE" then
				WindowStatus = "Complete"
			end if
		rs.movenext
		loop
		end if
		%>
		<li> <% response.write JobBarcode %> -  <% response.write WindowStatus %>

		<%
		rs2.movenext
		loop
		
		
		%>		
			
			
	</ul>
	
	<%
	end if
	
	rs.close
set rs=nothing


DBConnection.close
set DBConnection=nothing
	
	%>
	
	
</form>	


</body>
</html>



