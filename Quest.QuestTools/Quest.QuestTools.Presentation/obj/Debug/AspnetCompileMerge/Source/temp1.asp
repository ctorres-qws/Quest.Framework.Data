
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Glazing Statistics for Scanning based on Job and Floor-->
<!-- THis page shows Styles and Openings on Each Job And Floor-->
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
DIM AllScan(8)

SUcount = 0
SPcount = 0
SBcount = 0
Pcount = 0
Doorcount = 0
Ocount = 0

ScanSUcount = 0
ScanSPcount = 0
ScanSBcount = 0
ScanPcount = 0
ScanDoorcount = 0
ScanOcount = 0

Job = request.querystring("Job")
Floor = request.querystring("Floor")

	if Job = "" then	
		Job = "AAA"
	end if
	
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "Select * FROM " & JOB & " order by TAG ASC"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
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
		
	end if 

Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "Select * FROM Styles"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

if not rs2.eof then

Do while not rs2.eof
	rs.filter = " Job = '" & RS2("Job") & "' and Floor = '" & RS2("Floor") & "' and Tag = '" & RS2("Tag") & "'"
	WindowStyle = RS2("style")
	rs3.filter = "Name = '" & WindowStyle & "'" 
	x =1 
	do until x =9
		AllScan(x) = FALSE	'''  NEED TO FIX THIS need to read the Openings and the Style to check completed.
	x = x +1
	loop
	do while not rs.eof
		'Determine which openings are scanned (and ignore duplicates)
		x =1
		do until x =9
			if  RS3("O" & x) = "-" or RS3("O" & x) = "" then
			else 
				AllScan(x) = TRUE
			end if
		x = x +1
		loop
	rs.movenext 
	loop
	rs.filter =""
	'Now check this record scanned items against the style table and add count
	x = 1
	do until x =9
		if AllScan(X) = TRUE then
			if RS3("O" & x) = "SU" then
				SUcount = SUcount + 1
			end if
			if RS3("O" & x) = "SP" or RS3("O" & x) = "VP" then
				SPcount = SPcount + 1
			end if
			if RS3("O" & x) = "SB" or RS3("O" & x) = "VB" then
			SBcount = SBcount + 1
			end if
			if RS3("O" & x) = "P"  or RS3("O" & x) = "PP"  or RS3("O" & x) = "PF" or RS3("O" & x) = "FP" or RS3("O" & x) = "PE" or RS3("O" & x) = "VP" then
			Pcount = Pcount + 1
			end if
			if RS3("O" & x) = "OV" then
				OVcount = OVcount + 1
			end if
			if RS3("O" & x) = "SW" or RS3("O" & x) = "SD"  then
				Doorcount = Doorcount + 1
			end if
			if counted = FALSE then
				Othercount = OtherCount + 1 
			end if	
		end if
	x =x+1
	loop
	
	
	
rs2.movenext
loop

end if

	rs.close
set rs=nothing

	rs2.close
set rs2=nothing

	rs3.close
set rs3=nothing

DBConnection.close
set DBConnection=nothing
	


%>  


</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>
<form id="screen1" title="Glaze Stats" class="panel" name="Glaze" action="GlazingOpenings.asp" method="GET" selected="true">
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

	
	<%
	If JobFloor = TRUE then
	
	%>
	
		<ul id="screen2" title="Windows" selected="true">

		<li class="group">All Windows for <%response.write JOB & " " & Floor%> </li>
		
	
	<li>Sealed Unit (SU) - <%response.write SUcount%></li>
	<li>Panel (Panel) - <%response.write Pcount%></li>
	<li>Spandrel (SP) - <%response.write SPcount%></li>
	<li>Awning (OV) - <%response.write OVcount%></li>
	<li>Doors (Door) - <%response.write Doorcount%></li>
	<li>ShadowBox (SB) - <%response.write SBcount%></li>
	<li>Other (Other) - <%response.write Ocount%></li>
	<li>ErrorCount - <%response.write Errorcount%></li>
			
	</ul>
	
	<%
	end if
	

	%>
	
	
</form>	


</body>
</html>



