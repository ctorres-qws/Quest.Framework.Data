<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			
			<!--BARCODEemail.aspx - is an E-mail report of this information in webmail form, it must be updated when the code here gets changed-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<!-- Request from Jody Cash on January 9th 2014 to change the Auto Refresh from 1200 to 90 -->
  <meta http-equiv="refresh" content="90" >
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
	

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


' December 10 Logic change creates this report from a whole new method - see BARCODERTV- DECEMBER10 for Backup of old logic
' All of the RS2 functions were updated for Today's Activity Section


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_BARCODE_LINEITEM ORDER BY JOB ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection	

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection


Set rs5 = Server.CreateObject("adodb.recordset")
'strSQL5 = "SELECT * From X_EMPLOYEES"
strSQL5 = "SELECT * From X_BARCODEP ORDER BY DATETIME DESC"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection


Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEGA ORDER BY DATETIME DESC"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection

' RS7-10 Created later when they need to be activated in the report.

Set rs11 = Server.CreateObject("adodb.recordset")
strSQL11 = "SELECT * From X_BARCODEOV ORDER BY DATETIME DESC"
rs11.Cursortype = 2
rs11.Locktype = 3
rs11.Open strSQL11, DBConnection

JOB = UCASE(REQUEST.QueryString("JOB"))
FL = REQUEST.QueryString("FLOOR")

totalg = 0
totalg2 = 0
totalg2y = 0
totalgy = 0
totalay = 0
totala = 0
totalc = 0
'totalsu = 0
'totalsp = 0
' TotalSu and TotalSp are Glassline records in total
' FOREL and WILLIAN GlassLine are reported below and replace the TotalSP and TotalSU from above
totalsuForel = 0
totalspForel = 0
totalsuWillian = 0
totalspWillian = 0
'Panel Line and Awning LIne added to the Report - from X_BARCODEP/ X_BarcodeOV - New Table same as X_BARCODEGA
totalP = 0
totalOVA = 0
totalOVB = 0
totalOVC = 0
totalOVD = 0
totalOVAy = 0
totalOVBy = 0
totalOVCy = 0
totalOVDy = 0

%>

<!--#include file="todayandyesterday.asp"-->

<%

rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the length of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalg = totalg + 1
end if

IF rs("DEPT") = "GLAZING2" then
totalg2 = totalg2 + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totala = totala + 1
end if




  rs.movenext
loop

' Change to set up Friday stats show on MONDAY

if weekday(currentDate) = 2 then
monday = 1
rs.filter = "WEEK = " & lastweek & " AND YEAR = " & cYear & " AND DAY = " & Day((DateAdd("d",-2,currentDate)))
else
monday = 0
rs.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary
end if



Do while not rs.eof

'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalgy = totalgy + 1
end if

IF rs("DEPT") = "GLAZING2" then
totalg2y = totalg2y + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totalay = totalay + 1
end if

  rs.movenext
loop


rs6.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs6.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not

'GLASSLINE WAS REPLACED BY FOREL AND WILLIAN - These two are commented out and replaced by the two sets below
'IF rs6("DEPT") = "GLASSLINE" AND (rs6("TYPE") = "SU" OR rs6("TYPE") = "OV") then
'totalsu = totalsu + 1
'end if
'IF rs6("DEPT") = "GLASSLINE" AND rs6("TYPE") = "SP" then
'totalsp = totalsp + 1
'end if
 
 'Forel Line - of glass Scanned - Added March 10th 2014
IF UCASE(rs6("DEPT")) = "FOREL" AND (rs6("TYPE") = "SU" OR rs6("TYPE") = "OV") then
totalsuForel = totalsuForel + 1
end if

IF UCASE(rs6("DEPT")) = "FOREL" AND rs6("TYPE") = "SP" then
totalspForel = totalspForel + 1
end if


 'Willain Line - of glass Scanned  - Added March 10th 2014
 IF UCASE(rs6("DEPT")) = "WILLIAN" AND (rs6("TYPE") = "SU" OR rs6("TYPE") = "OV") then
totalsuWillian = totalsuWillian + 1
end if

IF UCASE(rs6("DEPT")) = "WILLIAN" AND rs6("TYPE") = "SP" then
totalspWillian = totalspWillian + 1
end if
 
 
 
  rs6.movenext
loop

rs5.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs5.eof
	
	IF rs5("DEPT") = "Panel" then
		totalP = totalP + 1
	end if
rs5.movenext
loop
'
'rs11.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
'Do while not rs11.eof
	
'	IF rs11("DEPT") = "FrameAssemble" then
'		totalOVA = totalOVA + 1
'	end if
'	IF rs11("DEPT") = "SashAssemble" then
'		totalOVB = totalOVB + 1
'	end if
'	IF rs11("DEPT") = "SashGlaze" then
'		totalOVC = totalOVC + 1
'	end if
'	IF rs11("DEPT") = "WindowMount" then
'		totalOVD = totalOVD + 1
'	end if
'rs11.movenext
'loop

%>
</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Glass Stats</li>
		<li><% response.write "GLAZING: " & totalg %></li>
        <li><% response.write "GLAZING2: " & totalg2 %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "PANEL: " & totalP %></li>
<!--		<li class="group">Awning Stats - In Development</li>
		<li><% response.write "Frame Assembly: " & totalOVA %></li>
		<li><% response.write "Sash Assembly: " & totalOVB %></li>
		<li><% response.write "Sash Glazing: " & totalOVC %></li>
		<li><% response.write "Window Mounting: " & totalOVD %></li>
-->		<li class="group">Glassline Split into Forel and Willian</li>
        <li><% response.write "FOREL: " & totalsuForel %></li>
        <li><% response.write "FOREL SP: " & totalspForel %></li>		
        <li><% response.write "WILLIAN: " & totalsuWillian %></li>
        <li><% response.write "WILLIAN SP: " & totalspWillian %></li>		
		
        
		<% 
		if monday = 1 then
		%>
			<li class="group">Saturday's Stats</li>
		<%
		else
		%>
			<li class="group">Yesterday's Stats</li>
		
		<%
		end if
		%>
		
		<li><% response.write "GLAZING: " & totalgy %></li>
        <li><% response.write "GLAZING2: " & totalg2y %></li>
		<li><% response.write "ASSEMBLY: " & totalay %></li>
                <li class="group">Today's Activity</li>
        


<%
	DIM JFCHECKID, wcount
'Declared Variables for the Today's Activities section
wcount=0
' Reset the filter on X_Barcode back to full (in the previous section, it was filtered to yesterday)
rs.filter = ""

'THIS IS SLOW, BUT IT MAKES SURE THAT ALL WINDOWS RELATED TO A BATCH ARE COUNTED FROM MORE THEN TODAY'S LIST
rs.movefirst 

do while not rs.eof
'Filter out RS2 items that do not have an ID
rs2.filter = "ID > 0 "
' Count =1 Means already counted from Barcode, so skip,  Any other value is to be counted here and then COUNT SET to 1 
	IF RS("COUNT") <> 1 Then 

		
		JFCHECKID = 0
					do while not rs2.eof
	

						IF UCASE(rs2("JOB")) = UCASE(rs("JOB")) AND rs2("FLOOR") = rs("Floor") AND rs2("DEPT") = rs("DEPT") THEN
							JFCHECKID = RS2("ID")
							wcount = wcount + 1	
						END IF
						rs2.movenext
					LOOP
			
		
					IF JFCHECKID = "0" THEN
						rs2.addnew 
						rs2.fields("JOB") = UCASE(RS("JOB"))
						rs2.fields("FLOOR") = RS("FLOOR")
						rs2.fields("DEPT") = RS("DEPT")
						rs2.fields("YEAR") = RS("YEAR")
						rs2.fields("MONTH") = RS("MONTH")
						rs2.fields("DAY") = RS("DAY")
						rs2.fields("WEEK") = RS("WEEK")
						rs2.fields("TAG") = 1
					
						rs2.UPDATE
					ELSE	
					
						if JFCHECKID = "" then
							response.write "<li>GOTCHA NULL</li>"
						else
							rs2.filter = "ID = " & JFCHECKID
							rs2.fields("TAG") = rs2.Fields("TAG") + wcount
							
							
						'If rs.fields("DAY") > rs2.fields("TODAYDAY") then
							'todaycount = 0
							
							'if rs DAy = cDay then
							'rs2.fields("TODAYTOTAL") = rs2(TODAYTOTAL) + todaycount
							' end if
							
							'end if
							
							' Updates the Date in the X_BARCODE_LINEITEM table to match the most current date 
							' Do I want this to be   rs2.fields("YEAR") = cYear    rs2.fields("YEAR") = RS("YEAR")
							
							rs2.fields("YEAR") = RS("YEAR")
							rs2.fields("MONTH") = RS("MONTH")
							rs2.fields("DAY") = RS("DAY")
							rs2.fields("WEEK") = RS("WEEK")
							
							'rs2.fields("YEAR") = cYear
							'rs2.fields("MONTH") = cMonth
							'rs2.fields("DAY") = cDay
							'rs2.fields("WEEK") = weeknumber	
							
							rs2.update
						end if 
					end if
					
				
				wcount = 0
	' Flags the Count in Barcode to 1 so that it is not counted again			
	RS("COUNT") = 1
	End If
	rs.movenext
loop	

' Today Count for each job
rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
rs2.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
rs2.movefirst 
do while not rs2.eof
'Filter out RS2 items that do not have an ID

	todayjob = 0
	
	rs.movefirst 
	do while not rs.eof
		if rs("JOB") = rs2("JOB") AND rs("Floor") = rs2("FLOOR") AND rs("DEPT") = rs2("DEPT") then
		todayjob = todayjob + 1
		end if
	rs.movenext
	loop
	rs2("LAST") = todayjob
rs2.movenext
loop
rs.filter = ""
rs2.filter = ""



	'Filters all the information that is in X_BARCODE_LINEITEM to just show items worked on Today
	rs2.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear 
	do while not rs2.eof
	rs4.filter = "Job = '" &  UCASE(rs2("Job")) & "' AND Floor = '" &  rs2("Floor") & "'"
	if rs4.bof then
	else
		IF rs2("DEPT") = "ASSEMBLY" then
			rs4.fields("DoneWinA") = rs2.fields("Tag")
		end if
		IF rs2("DEPT") = "GLAZING" then
			rs4.fields("DoneWinG") = rs2.fields("Tag")
		end if
		IF rs2("DEPT") = "GLAZING2" then
			rs4.fields("DoneWinG2") = rs2.fields("Tag")
		end if
		if NOT rs2("DEPT") = "GLAZING2" then
			'  Jody Reuqested that Glazing2 be removed from this section June 2014
			' Not filtered above so that X_Barcode_lineitem still collects the data, only the showing is removed.
			response.write "<li>" & rs2("DEPT") & " " & UCASE(rs2("Job")) & " " & rs2("Floor") & ": (" & rs2("Last") & ") " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
		end if
	
	
	'Insert in here a loop through rs4 that filters the Job / Floor and updates the "DoneWin" field with rs2("Tag") which is really the count of scanned windows
	end if
	rs2.movenext
	
	loop
' Close RS2 as soon as it is no longer in use.	
rs2.close
set rs2=nothing
	
	
	
	%>
	
	
	
<%
	
	
	
	
	
	
	
' Commented out, does not appear to link to any active commands, December 9th, Michael bernholtz	
' Added back to the report May 2014 - Slava Kotok and Michael Bernholtz fixed the error on the table
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
	RESPONSE.WRITE "<li>" & rs7("jobNumber") & " " & rs7("Cutstatus") & "%</li>"
	'  May 2014 - Working -  RESPONSE.WRITE "<LI>N/A</LI>"
	RESPONSE.WRITE "<li class='group'>EW Jamb Machine</li>"
	RESPONSE.WRITE "<li>" & rs8("jobNumber") & " " & rs8("Cutstatus") & "%</li>"
	
		RESPONSE.WRITE "<li class='group'>Q Width Machine</li>"
	RESPONSE.WRITE "<li>" & rs9("jobNumber") & " " & rs9("Cutstatus") & "%</li>"
	
		RESPONSE.WRITE "<li class='group'>Q Jamb Machine</li>"
	RESPONSE.WRITE "<li>" & rs10("jobNumber") & " " & rs10("Cutstatus") & "%</li>"


	%> 
 

        </ul>
        
  
<% 



rs.close
set rs=nothing
'rs2 closed above
'RS3 No longer in use, so it is commented out
'rs3.close
'set rs3=nothing
rs4.close '(causing error, Unknown so Commented out)
set rs4 = nothing
rs5.close
set rs5=nothing
rs6.close
set rs6=nothing
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
