<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 <!-- March 13 Backup is old version, Lev Requested new format, This page shows by job and departments are all in one line. -->
		 <!--Cleaned up old and unused portions of the code that have just been commented out -->
		 <!-- Cutting Machine information has been removed, it is duplicate information from Window Production in BarcoderTv -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1000" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript"> iui.animOn = true; </script>
  <!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>
<!-- Fixed Headers -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/extensions/FixedHeader/js/dataTables.fixedHeader.js"></script>
 
  <script type="text/javascript">
  $(document).ready( function () {
    $('#JobView').DataTable();
} );
  
  </script>
  <style>
body{
zoom: 90%
};
 </style>

<% 

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * From X_BARCODE_LINEITEM"
rs2.CursorLocation = 3
rs2.Cursortype = 3
rs2.Locktype = 2
rs2.Open strSQL2, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD ORDER BY JOB, FLOOR"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_BARCODEGA"
rs5.CursorLocation = 3
rs5.Cursortype = 3
rs5.Locktype = 2
rs5.Open strSQL5, DBConnection


sixweeks = weekNumber - 6
%>
<!--#include file="todayandyesterday.asp"-->


</head>

<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


	              <li class="group">LAST 6 WEEK'S ACTIVITY</li>
        
<%


Dim  RecentDate
Dim JambAverage
Dim WidthAverage


' Donewin(A,G,G2) contain the Assembly, Glazing, Glazing2 Counts in X_WIN_PROD
' HCUTC#  contain the Width Machine Percentages in X_WIN_PROD	
' CUTC#  contain the Jamb Machine Percentages in X_WIN_PROD	
	
	
	
	rs2.movefirst
	rs2.filter = "WEEK > " & sixweeks & " AND YEAR = " & cYear
	rs2.sort = "JOB, FLOOR, DEPT"
	
	
	rs4.movefirst
	
	' 6 weeks Activity is made up in the last 42 days
	Date2 = Date()-42
	
	rs4.filter = "DATESTAMP > #" & Date2 & "#"
	
	Response.write "<p> <table border ='1' class='JobView' id='JobView' cellpadding='1' >"
	Response.write "<thead><tr title = 'Click on a Header to Sort by that Column' ><th>Job</th><th>Floor</th><th>Date</th><th>Glazed%<th>Cycles</th><th>Total Windows</th><th>SqFt Total</th><th>Assembly</th><th>Date</th><th>Glazing</th><th>Date</th><th>Glazing2</th><th>Date</th><th>Forel</th><th>Willian</th></thead><tbody>"
	'Response.write "<th>Width C1</th><th>Jamb C1</th><th>Width C2</th><th>Jamb C2</th><th>Width C3</th><th>Jamb C3</th>"
	'Response.write "<th>Width C4</th><th>Jamb C4</th><th>Width C5</th><th>Jamb C5</th><th>Width C6</th><th>Jamb C6</th><th>Width C7</th><th>Jamb C7</th><th>Width C8</th><th>Jamb C8</th><th>Width C9</th><th>Jamb C9</th><th>Width C10</th><th>Jamb C10</th></tr>"
	
	do while not rs4.eof
	
	
	
	' Write all the details to the table
		response.write "<tr><td>" & rs4("Job") & "</td><td>" & rs4("Floor") & "</td><td>" & rs4("Datestamp") &" </td>"
%> <!--
Lev was unhappy with the way this was showing so it is being removed.
This would be filled in from BarcoderTV (That method was not optimal and has been deleted, better to update and then report
If a newer method is developed, then fill it in and return. - Michael Bernholtz, (permission, Jody Cash) April 2014
		Jobcycles = rs4("cycles")
		Select Case Jobcycles
			Case 1
				JambAverage =  rs4("HCUTC1")
				WidthAverage = rs4("CUTC1")
			Case 2
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2"))/2
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2"))/2
			Case 3
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3"))/3
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3"))/3
			Case 4
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4"))/4
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4"))/4
			Case 5
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5"))/5
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5"))/5
			Case 6
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5")+rs4("HCUTC6"))/6
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5")+rs4("CUTC6"))/6
			Case 7
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5")+rs4("HCUTC6")+rs4("HCUTC7"))/7
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5")+rs4("CUTC6")+rs4("CUTC7"))/7
			Case 8
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5")+rs4("HCUTC6")+rs4("HCUTC7")+rs4("HCUTC8"))/8
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5")+rs4("CUTC6")+rs4("CUTC7")+rs4("CUTC8"))/8
			Case 9
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5")+rs4("HCUTC6")+rs4("HCUTC7")+rs4("HCUTC8")+rs4("HCUTC9"))/9
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5")+rs4("CUTC6")+rs4("CUTC7")+rs4("CUTC8")+rs4("CUTC9"))/9
			Case 10
				JambAverage =  (rs4("HCUTC1")+rs4("HCUTC2")+rs4("HCUTC3")+rs4("HCUTC4")+rs4("HCUTC5")+rs4("HCUTC6")+rs4("HCUTC7")+rs4("HCUTC8")+rs4("HCUTC9")+rs4("HCUTC10"))/10
				WidthAverage = (rs4("CUTC1")+rs4("CUTC2")+rs4("CUTC3")+rs4("CUTC4")+rs4("CUTC5")+rs4("CUTC6")+rs4("CUTC7")+rs4("CUTC8")+rs4("CUTC9")+rs4("CUTC10"))/10
			Case Else
				JambAverage =  rs4("HCUTC1")
				WidthAverage = rs4("CUTC1")	
		End Select
				if not isnull(JambAverage) then 
					JambAverage = Round(CINT(JambAverage))
					JambAverage = JambAverage & "%"
				end if
				if not isnull(WidthAverage)  then
					WidthAverage = Round(CINT(WidthAverage))
					WidthAverage = WidthAverage & "%"
				end if
		
		response.write "<td>" & JambAverage & "</td><td>" & widthAverage & "</td>"
-->
<%			
		' Percentage Glazed of Total
	
			glazedAverage = rs4("DonewinG")/rs4("TotalWin")*100

				if not isnull(glazedAverage)  then
					glazedAverage = Round(CINT(glazedAverage))
					glazedAverage = glazedAverage & "%"
				end if
	
			response.write "<td>" & glazedAverage & "</td>"	
		
		
		
		
		
		response.write "<td>" & rs4("cycles") & "</td>" 
		response.write "<td>" & rs4("TotalWin") & "</td>"
		response.write "<td>" & rs4("TotalSqFt") &" ft<sup>2</sup> </td>"
		

	

		response.write "<td>" & rs4("DonewinA") & "</td>" 

		' Get the Date from the BARCODE LINE ITEM - Assembly
		
				rs2.filter = "Job = '" &  rs4("Job") & "' AND Floor = '" &  rs4("Floor") & "' AND DEPT = 'ASSEMBLY'"
		
				if rs2.bof then
					response.write "<td></td>"
				else
	
					RecentDate = DateValue(rs2("month") & "-" & rs2("day") & "-" & rs2("year"))
		
					response.write "<td>" & RecentDate &"</td>"

				end if
		
		response.write "<td>" & rs4("DonewinG") & "</td>"
		
		' Get the Date from the BARCODE LINE ITEM - Glazing
		
				rs2.filter = "Job = '" &  rs4("Job") & "' AND Floor = '" &  rs4("Floor") & "' AND DEPT = 'GLAZING'"
		
				if rs2.bof then
					response.write "<td></td>"
				else
	
					RecentDate = DateValue(rs2("month") & "-" & rs2("day") & "-" & rs2("year"))
		
					response.write "<td>" & RecentDate &"</td>"

				end if
		
		response.write "<td>" & rs4("DonewinG2") & "</td>"
		
		' Get the Date from the BARCODE LINE ITEM - Glazing
		
				rs2.filter = "Job = '" &  rs4("Job") & "' AND Floor = '" &  rs4("Floor") & "' AND DEPT = 'GLAZING2'"
		
				if rs2.bof then
					response.write "<td></td>"
				else
	
					RecentDate = DateValue(rs2("month") & "-" & rs2("day") & "-" & rs2("year"))
		
					response.write "<td>" & RecentDate &"</td>"

				end if	

		' HCUT and CUT
		
		'response.write "<td>" & rs4("HCUTC1") & "%</td><td>" & rs4("CUTC1") & "%</td><td>" & rs4("HCUTC2") & "%</td><td>" & rs4("CUTC2") & "%</td><td>" & rs4("HCUTC3") & "%</td><td>" & rs4("CUTC3") & "%</td>"		
		'response.write "<td>" & rs4("HCUTC4") & "%</td><td>" & rs4("CUTC4") & "%</td><td>" & rs4("HCUTC5") & "%</td><td>" & rs4("CUTC5") & "%</td><td>" & rs4("HCUTC6") & "%</td><td>" & rs4("CUTC6") & "%</td><td>" & rs4("HCUTC7") & "%</td><td>" & rs4("CUTC7") & "%</td><td>" & rs4("HCUTC8") & "%</td><td>" & rs4("CUTC8") & "%</td><td>" & rs4("HCUTC9") & "%</td><td>" & rs4("CUTC9") & "%</td><td>" & rs4("HCUTC10") & "%</td><td>" & rs4("CUTC10") & "%</td>"		
		
		
		
	'Added Portion to Show Willian and Forel for Job and Floor	
		rs5.filter = "DATETIME > #" & Date2 & "# AND JOB = '" & rs4("JOB") & "' AND FLOOR = '" & rs4("Floor") & "' AND DEPT = 'Forel' "
		ForelCount = rs5.recordcount
		rs5.filter = ""
		rs5.filter = "DATETIME > #" & Date2 & "# AND JOB = '" & rs4("JOB") & "' AND FLOOR = '" & rs4("Floor") & "' AND DEPT = 'Willian' "
		WillianCount = rs5.recordcount
		rs5.filter = ""
		
		response.write "<td>" & ForelCount &"</td>"
		response.write "<td>" & WillianCount &"</td>"
		
		
		response.write "</tr>"
		
		
		
		
	rs4.movenext
	loop
	
	
Response.write "</tbody></table></p>"
rs2.close
set rs2=nothing
rs4.close
set rs4=nothing
rs5.close
set rs5=nothing	

Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "SELECT * From PROECOHOR ORDER BY ID DESC"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection

Set rs8 = Server.CreateObject("adodb.recordset")
strSQL8 = "SELECT * From PROECOVERT ORDER BY ID DESC"
rs8.Cursortype = 2
rs8.Locktype = 3
rs8.Open strSQL8, DBConnection

Set rs9 = Server.CreateObject("adodb.recordset")
strSQL9 = "SELECT * From PROQHOR ORDER BY ID DESC"
rs9.Cursortype = 2
rs9.Locktype = 3
rs9.Open strSQL9, DBConnection

Set rs10 = Server.CreateObject("adodb.recordset")
strSQL10 = "SELECT * From PROQVERT ORDER BY ID DESC"
rs10.Cursortype = 2
rs10.Locktype = 3
rs10.Open strSQL10, DBConnection

	
	RESPONSE.WRITE "<li class='group'>EW Width Machine</li>"
	RESPONSE.WRITE "<li>" & rs7("jobNumber") & ": " & rs7("Cutstatus") & "%</li>"
	'RESPONSE.WRITE "<LI>N/A</LI>"
	RESPONSE.WRITE "<li class='group'>EW Jamb Machine</li>"
	RESPONSE.WRITE "<li>" & rs8("jobNumber") & ": " & rs8("Cutstatus") & "%</li>"
	
		RESPONSE.WRITE "<li class='group'>Q Width Machine</li>"
	RESPONSE.WRITE "<li>" & rs9("jobNumber") & ": " & rs9("Cutstatus") & "%</li>"
	
		RESPONSE.WRITE "<li class='group'>Q Jamb Machine</li>"
	RESPONSE.WRITE "<li>" & rs10("jobNumber") & ": " & rs10("Cutstatus") & "%</li>"
	
	
	
	
	
	%> 
 

        </ul>
        
        
     <% 


rs7.close
set rs7=nothing
rs8.close
set rs8=nothing
rs9.close
set rs9=nothing
rs10.close
set rs10=nothing
DBConnection.close
set DBConnection=nothing
%> 





</body>
</html>



