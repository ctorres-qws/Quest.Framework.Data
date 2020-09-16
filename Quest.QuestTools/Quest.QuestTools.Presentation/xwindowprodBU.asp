<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		 <!-- This code was Deleted by a data connection loss and rewritten on February 26th 2015 -->
		 <!-- Job View Shows all Jobs Worked on in the last 6 weeks - Broken down into Assembly / Glazing / Glazing2-->
		 
		 
		 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1200" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript"> iui.animOn = true; </script>
  <script src="sorttable.js"></script>
<!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

  <script type="text/javascript">
  $(document).ready( function () {
    $('#Job').DataTable();
} );
  
  </script>
  
  
</head>

<body>
<!--#include file="todayandyesterday.asp"-->


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

	              <li class="group">LAST 6 WEEK'S ACTIVITY</li>
				  
				 <li><table border='1' class='Job' id ='Job' ><thead><tr><th>Job</th><th>Floor</th><th>SqFt</th><th>Total Window</th><th>Assembled</th><th>Date</th><th>Glazed</th><th>Date</th><th>Glaze 2</th><th>Date</th><th>Forel</th><th>Date</th><th>Willian</th><th>Date</th></tr></thead><tbody>

<%		
SixWeek = Date()-42		  
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT DateStamp, JOB, Floor,TotalWin, TotalSqFt From X_WIN_PROD where DATESTAMP > #" & SixWeek & "# order by DATESTAMP DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JOB") & " </td>"
	response.write "<td>" & RS("Floor") & " </td>"
	response.write "<td>" & RS("TotalSqFT") & " ft<sup>2</sup></td>"
	response.write "<td>" & RS("TotalWin") & " </td>"
	
	' Now find each floor in X_Barcode
	
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT JOB, Floor, DEPT, DateTime From X_BARCODE where JOB = '" & RS("JOB") & "' and Floor = '" & RS("Floor") & "' ORDER BY DEPT"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL2, DBConnection
	Assemble = 0
	AssembleDate = #01/01/1999#
	Glaze = 0
	GlazeDate = #01/01/1999#
	Glaze2 = 0
	Glaze2Date = #01/01/1999#
	
	do while not rs2.eof
		select case rs2("Dept")
		case "ASSEMBLY"
			Assemble = Assemble + 1
			if rs2("DateTime") > AssembleDate then
				AssembleDate = rs2("DateTime")
			end if
		case "GLAZING"
			Glaze = Glaze + 1
			if rs2("DateTime") > GlazeDate then
				GlazeDate = rs2("DateTime")
			end if
		case "GLAZING2"
			Glaze2 = Glaze2 + 1
			if rs2("DateTime") > Glaze2Date then
				Glaze2Date = rs2("DateTime")
			end if
		end select
	rs2.movenext
	loop
	rs2.close
	set rs2 = nothing
	
	response.write "<td>" & Assemble & " </td>"
	if AssembleDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(AssembleDate,2) & " </td>"
	end if
	
	response.write "<td>" & Glaze & " </td>"
	if GlazeDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(GlazeDate,2) & " </td>"
	end if
	response.write "<td>" & Glaze2 & " </td>"
	if Glaze2Date = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(Glaze2Date,2) & " </td>"
	end if	
	
		' Now find each floor in X_BarcodeGA for Window Sealing
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT  JOB, Floor, DEPT, DateTime From X_BARCODEGA where JOB = '" & RS("JOB") & "' and FLoor = '" & RS("Floor") & "' order by Dept"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL3, DBConnection
	Forel = 0
	ForelDate = #01/01/1999#
	Willian = 0
	WillianDate = #01/01/1999#
	
	do while not rs3.eof
		select case rs3("Dept")
			case "Forel"
				Forel = Forel + 1
				if rs3("DateTime") > ForelDate then
					ForelDate = rs3("DateTime")
				end if
			case "Willian"
				Willian = Willian + 1
				if rs3("DateTime") > WillianDate then
					WillianDate = rs3("DateTime")
				end if
			end select
		rs3.movenext
		loop
		rs3.close
		set rs3 = nothing
		
		
	response.write "<td>" & Forel & " </td>"
		if ForelDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(ForelDate,2) & " </td>"
	end if	
	response.write "<td>" & Willian & " </td>"
		if WillianDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(WillianDate,2) & " </td>"
	end if	
	response.write "</tr>"
	
rs.movenext
loop

rs.close
set rs=nothing
DBConnection.close
set DBConnection = nothing
	
%>
</tbody></table>
</ul>	

</body>
</html>



