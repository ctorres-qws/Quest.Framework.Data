<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Production Data to be sent to V_REPORT3 -->
<!-- Daily Production Data - Replacing BarcodeWeekly.asp-->
<!--Requested by Jody Cash - Implemented by Michael Bernholtz, July 15th -->
<!-- Jody Cash set up as Stored Procedure to run each Half hour -->
<!-- updated to sort firstcomplete descending which syncronized both v1 and v3 reports, Sept 1 2016-->
<!--Updated June 2015 for SQFT-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Add Data to Database</title>
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

'Create a Query
    SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * From X_EMPLOYEES"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

'rs4, rs5, rs6 - Glass, Panels, Awning - Below

' Today Values Total
totalg = 0
totala = 0
totalg2 = 0

' Today Values Day Shift
dayg = 0
daya = 0
dayg2 = 0

' Today Values Night Shift
nightg = 0
nighta = 0
nightg2 = 0


' Yesterday Values Total
totalgy = 0
totalay = 0
totalg2y = 0

' Yesterday Values Day Shift
daygy = 0
dayay = 0
dayg2y = 0

' Yesterday Values Night Shift
nightgy = 0
nightay = 0
nightg2y = 0

elist = ""

%>
<!--#include file="todayandyesterday.asp"-->
<%
rs.filter = "DAY = " & cDay & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear
Do while not rs.eof
	rs3.filter = "NUMBER = " & rs("EMPLOYEE")
	if rs3.bof then
		elist = elist & rs("EMPLOYEE") & "(" & rs("DEPT") & ") ,"
	else
		if rs3("Shift") = "0" then ' Day Shift

		SELECT CASE rs("DEPT")
				Case "GLAZING"
					dayg = dayg + 1
				Case "GLAZING2"
					dayg2 = dayg2 + 1
				Case "ASSEMBLY"
					daya = daya + 1
			End SELECT

		end if
		if rs3("Shift") = "1" then ' NIght Shift
			SELECT CASE rs("DEPT")
				Case "GLAZING"
					nightg = nightg + 1
				Case "GLAZING2"
					nightg2 = nightg2 + 1
				Case "ASSEMBLY"
					nighta = nighta + 1
			End SELECT
		end if
	end if



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

rs.filter = ""
rs.filter = "DAY = " & cYesterday & " AND WEEK = " & weekNumbery & " AND YEAR = " & cYeary
Do while not rs.eof
	' treat all yesterday as working despite counting the error, it still makes Glazing and SQFT values.
	if isnumeric(rs("Employee")) then
		rs3.filter = "NUMBER = " & rs("EMPLOYEE")
	else
		rs3.filter = "NUMBER = 0000"
	end if
	if rs3.bof then
		'elist = elist & rs("EMPLOYEE") & "(" & rs("DEPT") & ") ," 
		' THIS WAS REMOVED AS it is causing errors, 
		'Today's value was being filled in with yesterday's error
	else
		if rs3("Shift") = "0" then ' Day Shift

		SELECT CASE rs("DEPT")
				Case "GLAZING"
					daygy = daygy + 1
				Case "GLAZING2"
					dayg2y = dayg2y + 1
				Case "ASSEMBLY"
					dayay = dayay + 1
			End SELECT

		end if
		if rs3("Shift") = "1" then ' NIght Shift
			SELECT CASE rs("DEPT")
				Case "GLAZING"
					nightgy = nightgy + 1
				Case "GLAZING2"
					nightg2y = nightg2y + 1
				Case "ASSEMBLY"
					nightay = nightay + 1
			End SELECT
		end if
	end if



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

' Completed the Logic to fill the Values

rs.close
set rs=nothing
rs3.close
set rs3=nothing


GLASS_FOREL = 0
GLASS_WILLIAN = 0
GLASS_FORELy = 0
GLASS_WILLIANy = 0

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_BARCODEGA"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

rs4.filter = "DAY = " & cDay & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear
Do while not rs4.eof

		SELECT CASE rs4("DEPT")
				Case "Forel"
					GLASS_FOREL = GLASS_FOREL + 1
				Case "Willian"
					GLASS_WILLIAN = GLASS_WILLIAN + 1

			End SELECT
rs4.movenext
loop

rs4.filter = "DAY = " & cYesterday & " AND WEEK = " & weekNumbery & " AND YEAR = " & cYeary
Do while not rs4.eof

		SELECT CASE rs4("DEPT")
				Case "Forel"
					GLASS_FORELy = GLASS_FORELy + 1
				Case "Willian"
					GLASS_WILLIANy = GLASS_WILLIANy + 1

			End SELECT
rs4.movenext
loop


rs4.close
set rs4 = nothing

'rs5 ---------------------------------------------PANELS---------------------------------------------------

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_BARCODEP"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

totalP = 0
totalPy = 0

rs5.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs5.eof
	
	IF rs5("DEPT") = "Cut" or rs5("DEPT") = "Bend" or rs5("DEPT") = "Ship" or rs5("DEPT") = "Receive" then
		totalP = totalP + 1
	end if
rs5.movenext
loop


rs5.filter = ""
rs5.filter =" DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary 

Do while not rs5.eof
	
	IF rs5("DEPT") = "Cut" or rs5("DEPT") = "Bend" or rs5("DEPT") = "Ship" or rs5("DEPT") = "Receive" then
		totalPy = totalPy + 1
	end if
rs5.movenext
loop

rs5.close
set rs5 = nothing

'rs6 -----------------------------------------Awning (Glaze) --------------------------------------------------

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEOV"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection


totalOVA = 0
totalOVB= 0
totalOVC = 0
totalOVD = 0

totalOVAy = 0
totalOVBy = 0
totalOVCy = 0
totalOVDy = 0


rs6.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs6.eof
	
	IF rs6("DEPT") = "FrameAssemble" then
		totalOVA = totalOVA + 1
	end if
	IF rs6("DEPT") = "SashAssemble" then
		totalOVB = totalOVB + 1
	end if
	IF rs6("DEPT") = "SashGlaze" then
		totalOVC = totalOVC + 1
	end if
	IF rs6("DEPT") = "WindowMount" then
		totalOVD = totalOVD + 1
	end if
rs6.movenext
loop

rs6.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary
Do while not rs6.eof
	
	IF rs6("DEPT") = "FrameAssemble" then
		totalOVAy = totalOVAy + 1
	end if
	IF rs6("DEPT") = "SashAssemble" then
		totalOVBy = totalOVBy + 1
	end if
	IF rs6("DEPT") = "SashGlaze" then
		totalOVCy = totalOVCy + 1
	end if
	IF rs6("DEPT") = "WindowMount" then
		totalOVDy = totalOVDy + 1
	end if
rs6.movenext
loop


rs6.close
set rs6 = nothing




'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ZIPPER MACHINE CUTS
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

ZipperRed = 0
ZipperRedy = 0



Set rsRZ = Server.CreateObject("adodb.recordset")
strSQL = "Select * from ProZipperRED"
rsRZ.Cursortype = 2
rsRZ.Locktype = 3
rsRZ.Open strSQL, DBConnection	
	rsRZ.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear 
	
do while not rsRZ.eof
	ZipperRed = ZipperRed + 1
rsRZ.movenext
loop
	rsRZ.Filter = ""
	rsRZ.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary 
	
do while not rsRZ.eof
	ZipperRedy = ZipperRedy + 1
rsRZ.movenext
loop

rsRZ.close
set rsRZ= Nothing


ZipperBlue = 0
ZipperBluey = 0


Set rsBZ = Server.CreateObject("adodb.recordset")
strSQL = "Select * from ProZipperBlue"
rsBZ.Cursortype = 2
rsBZ.Locktype = 3
rsBZ.Open strSQL, DBConnection	
	rsBZ.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear 
	
do while not rsBZ.eof
	ZipperBlue = ZipperBlue + 1
rsBZ.movenext
loop
	rsBZ.Filter = ""
	rsBZ.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary 
	
do while not rsBZ.eof
	ZipperBluey = ZipperBluey + 1
rsBZ.movenext
loop

rsBZ.close
set rsBZ= Nothing


'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Glazing by Opening (New April 2016)
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "Select * FROM X_GLAZING order by BARCODE ASC, FIRSTCOMPLETE DESC"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection
rs7.filter = "DAY = " & cDay & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear


Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * From X_EMPLOYEES"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection


ScanCompleteWindow = 0
ScanWindow = 0
SCWDay = 0
SCWNight = 0

Do while not rs7.eof
rs3.filter = "NUMBER = " & rs7("EMPLOYEE")
	if rs3.bof then
		elist = elist & rs7("Employee") & ", " 
		ShiftEmployee = 0
	else
		ShiftEmployee = rs3("Shift")
	end if
		
		OldBarcode = Barcode
		Barcode = rs7("Barcode")
			AllScans = AllScans + 1	
			if OldBarcode = Barcode then
				if rs7("FirstComplete") = "TRUE" then
					ScanCompleteWindow = ScanCompleteWindow + 1
					ScanWindow = ScanWindow - 1
					if ShiftEmployee = "0" then ' Day Shift
						SCWDay = SCWDay + 1
					end if
					if ShiftEmployee = "1" then ' Night Shift
						SCWNight = SCWNight + 1
					end if
				end if
			else
				if rs7("FirstComplete") = "TRUE" then
				' ------------------------------Added Code March 17 2015 to collect Square footage based on Job Table -------------------
				JobName = rs7("job")
				FloorName = rs7("floor")
				TagName = rs7("Tag")

				Set rsSQFT = Server.CreateObject("adodb.recordset")
				strSQL = "Select X,Y FROM " & JobName & " Where JOB = '" &  JobName & "' and  Floor = '" &  FloorName & "' and Tag = '" &  TagName & "'"
				rsSQFT.Cursortype = 2 
				rsSQFT.Locktype = 3
				On Error Resume Next  
				rsSQFT.Open strSQL, DBConnection
				If rsSQFT.State = 1 Then 
					SquareFoot = rsSQFT("X") * rsSQFT("Y") / 144
					WindowPerimeter = (rsSQFT("X") * 2) + (rsSQFT("Y") * 2)
					rsSQFT.close
					set rsSQFT = nothing
				Else
					SquareFoot = 30
					WindowPerimeter = 30
				End if
				Response.write SquareFoot & "-"
				TotalSquareFoot = TotalSquareFoot + SquareFoot
				TotalWindowPerimeter = TotalWindowPerimeter + WindowPerimeter
				'--------------------------- End new Calculation ---------------------Code added at End for adding to Database ----------
				
					ScanCompleteWindow = ScanCompleteWindow + 1
					if ShiftEmployee = "0" then ' Day Shift
						SCWDay = SCWDay + 1
					end if
					if ShiftEmployee = "1" then ' Night Shift
						SCWNight = SCWNight + 1
					end if
				else
					ScanWindow = ScanWindow + 1
				end if
			end if

rs3.filter = ""	
rs7.movenext
loop

' Now Yesterday Values
rs7.filter = ""
rs7.filter = "DAY = " & cYesterday & " AND WEEK = " & weekNumbery & " AND YEAR = " & cYeary


ScanCompleteWindowy = 0
ScanWindowy = 0
SCWDayy = 0
SCWNighty = 0

Do while not rs7.eof
rs3.filter = "NUMBER = " & rs7("EMPLOYEE")
	if rs3.bof then
		ShiftEmployee = 0
	else
		ShiftEmployee = rs3("Shift")
	end if
OldBarcode = Barcode
Barcode = rs7("Barcode")
	if OldBarcode = Barcode then
		if rs7("FirstComplete") = "TRUE" then
			ScanCompleteWindowy = ScanCompleteWindowy + 1
			ScanWindowy = ScanWindowy - 1
			if ShiftEmployee = "0" then ' Day Shift
				SCWDayy = SCWDayy + 1
			end if
			if ShiftEmployee = "1" then ' Night Shift
				SCWNighty = SCWNighty + 1
			end if
		end if
	else
		if rs7("FirstComplete") = "TRUE" then
				' ------------------------------Added Code March 17 2015 to collect Square footage based on Job Table -------------------
				JobName = rs7("job")
				FloorName = rs7("floor")
				TagName = rs7("Tag")

				Set rsSQFT = Server.CreateObject("adodb.recordset")
				strSQL = "Select X,Y FROM " & JobName & " Where JOB = '" &  JobName & "' and  Floor = '" &  FloorName & "' and Tag = '" &  TagName & "'"
				rsSQFT.Cursortype = 2 
				rsSQFT.Locktype = 3
				On Error Resume Next  
				rsSQFT.Open strSQL, DBConnection
				If rsSQFT.State = 1 Then 
					SquareFoot = rsSQFT("X") * rsSQFT("Y") / 144
					WindowPerimeter =  (rsSQFT("X") * 2) + (rsSQFT("Y") * 2) 
					rsSQFT.close
					set rsSQFT = nothing
				Else
					SquareFoot = 30
					WindowPerimeter = 30
				End if
				TotalSquareFooty = TotalSquareFooty + SquareFoot
				TotalWindowPerimetery = TotalWindowPerimetery + WindowPerimeter

				
				'--------------------------- End new Calculation ---------------------Code added at End for adding to Database ----------
		
			ScanCompleteWindowy = ScanCompleteWindowy + 1
			if ShiftEmployee = "0" then ' Day Shift
				SCWDayy = SCWDayy + 1
			end if
			if ShiftEmployee = "1" then ' Night Shift
				SCWNighty = SCWNighty + 1
			end if
		else
			ScanWindowy = ScanWindowy + 1
		end if
	end if

rs3.filter =""
rs7.movenext
loop

rs7.close
set rs7 = nothing



'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ADD Values to V_REPORT3
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------




Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM V_REPORT3 ORDER BY ID DESC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection	


rs2.filter = "DAY = " & cDay & " AND WEEK = " & weekNumber & " AND YEAR = " & cYear

if rs2.eof then
	rs2.addnew
end if
rs2("DAY") = cDay
rs2("WEEK") = weekNumber
rs2("YEAR") = cYear
rs2("GLAZING") = totalg
rs2("GLAZING2") = totalg2
rs2("ASSEMBLY") = totala
rs2("DAY_GLAZING") = dayg
rs2("DAY_GLAZING2") = dayg2
rs2("DAY_ASSEMBLY") = daya
rs2("NIGHT_GLAZING") = nightg
rs2("NIGHT_GLAZING2") = nightg2
rs2("NIGHT_ASSEMBLY") = nighta
rs2("ERROR_EMPLOYEE") = elist
rs2("GLASS_FOREL") = GLASS_FOREL
rs2("GLASS_WILLIAN") = GLASS_WILLIAN
rs2("ZipperRed") = ZipperRed
rs2("ZipperBlue") = ZipperBlue
rs2("SquareFoot") = TotalSquareFoot
rs2("WindowPerimeter") = TotalWindowPerimeter
rs2("GlazingFull") = ScanCompleteWindow
rs2("GlazingPartial") = ScanWindow
rs2("GlazingFullD") = SCWDay
rs2("GlazingFullN") = SCWNight
rs2("Panel") = TotalP
rs2("Awning") = TotalOVC

rs2.update

rs2.filter = ""
rs2.filter = "DAY = " & cYesterday & " AND WEEK = " & weekNumbery & " AND YEAR = " & cYeary

if rs2.eof then
	rs2.addnew
end if
rs2("DAY") = cYesterday
rs2("WEEK") = weekNumbery
rs2("YEAR") = cYeary
rs2("GLAZING") = totalgy
rs2("GLAZING2") = totalg2y
rs2("ASSEMBLY") = totalay
rs2("DAY_GLAZING") = daygy
rs2("DAY_GLAZING2") = dayg2y
rs2("DAY_ASSEMBLY") = dayay
rs2("NIGHT_GLAZING") = nightgy
rs2("NIGHT_GLAZING2") = nightg2y
rs2("NIGHT_ASSEMBLY") = nightay
rs2("GLASS_FOREL") = GLASS_FORELy
rs2("GLASS_WILLIAN") = GLASS_WILLIANy
rs2("ZipperRed") = ZipperRedy
rs2("ZipperBlue") = ZipperBluey
rs2("SquareFoot") = TotalSquareFooty
rs2("WindowPerimeter") = TotalWindowPerimetery
rs2("GlazingFull") = ScanCompleteWindowy
rs2("GlazingPartial") = ScanWindowy
rs2("GlazingFullD") = SCWDayy
rs2("GlazingFullN") = SCWNighty
rs2("Panel") = TotalPy
rs2("Awning") = TotalOVCy
rs2.update

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing


 ' Display Results as Confirmation
  
 
%>
</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Today Total Stats</li>
		<li><% response.write "GLAZING Full: " & ScanCompleteWindow & " - " & TotalSquareFoot & "ft<sup>2</sup>"%></li>
		<li><% response.write "GLAZING Partial: " & ScanWindow %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "GLASSLINE - FOREL: " & GLASS_FOREL %></li>
		<li><% response.write "GLASSLINE - WILLAIN: " & GLASS_WILLIAN %></li>
		<li><% response.write "ZIPPER - RED: " & ZipperRed %></li>
		<li><% response.write "ZIPPER - BLUE: " & ZipperBlue %></li>
		<li><% response.write "Panel: " & TotalP %></li>
		<li><% response.write "Awning: " & TotalOVC %></li>
       		<li class="group">Day Shift Stats</li>
		<li><% response.write "GLAZING Full Today: " & SCWDay %></li>
		<li><% response.write "GLAZING Full Yesterday: " & SCWDayy %></li>
		<li><% response.write "ASSEMBLY: " & daya %></li>
           		<li class="group">Night Shift Stats</li>
		<li><% response.write "GLAZING Full Today: " & SCWNight %></li>
		<li><% response.write "GLAZING Full Yesterday: " & SCWNighty %></li>
		<li><% response.write "ASSEMBLY: " & nighta %></li>
				<li class="group">Employee Errors</li>
		<li><% response.write "Invalid Employee Number(s): " & elist %></li>
       
	   </ul>

</body>
</html>



