<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--This is a Stored Procedure that runs every half an hour -->
			<!-- Fills V_Report1 and V_Report2 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <!--#include file="dbpath.asp"-->
<% Server.ScriptTimeout = 500 %> 
</head>
<body >

<%
' --------------------------------------------------------------------------------------------------Today
	currentDate = Date
	cDay = Day(currentDate)
	cMonth = Month(currentDate)
	cYear = Year(currentDate )

' --------------------------------------------------------------------------------------------------Yesterday (or Saturday on a Monday)	
	if weekday(currentDate) = 2 then
		cYesterday = Day((DateAdd("d",-2,currentDate)))
		cMonthy = Month((DateAdd("d",-2,currentDate)))
		cYeary = Year((DateAdd("d",-2,currentDate)))
		yesterdayDate = DateAdd("d", -2, Date)
	else
		cYesterday = Day((DateAdd("d",-1,currentDate)))
		cMonthy = Month((DateAdd("d",-1,currentDate)))
		cYeary = Year((DateAdd("d",-1,currentDate)))
		yesterdayDate = DateAdd("d", -1, Date)
	end if
	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' First set of Todays Data - All the Today and Yestderday individual Number
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------Collect Glazing and Assembly Data 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE WHERE DEPT = 'ASSEMBLY' ORDER BY DATETIME DESC"
' WHERE (DAY = " & cDAY & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear & ") OR  (DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary & ")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

' ----------------------------------------------------------------------------------------------Today's Glazing, Glazing2, And Assembly

totala = 0

rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs.eof
	totala = totala + 1
	rs.movenext
loop

' -----------------------------------------------------------------------------------------------Yesterday's (or Saturday's) Glazing, Glazing2, or Assembly 
' This report stores every day, so Yesterday information is now collected by the report, not as a data field

totalay = 0


rs.filter =""
rs.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary
Do while not rs.eof
	totalay = totalay + 1
	rs.movenext
loop

' -------------------------------------------------------------------------------------------------Panels Today 
totalP = 0
totalPy = 0

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_BARCODEP WHERE (DAY = " & cDAY & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear & ") OR  (DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary & ") ORDER BY DATETIME DESC"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

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
Response.write "Panels Today <br/>"
' -------------------------------------------------------------------------------------------------Glassline Today

totalsuForel = 0
'totalspForel = 0
totalsuWillian = 0
'totalspWillian = 0
totalOtherForel = 0
totalOtherWillian = 0
totalWillian = 0
totalForel= 0

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEGA WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " ORDER BY DATETIME DESC"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection

Do while not rs6.eof

 
 ' -------------------------------------------------------------------------------Forel Line - of glass Scanned - Added March 10th 2014
IF UCASE(rs6("DEPT")) = "FOREL" then
	IF rs6("TYPE") = "SU" OR rs6("TYPE") = "OV"  then
		totalsuForel = totalsuForel + 1
	else 
		totalOtherForel = totalOtherForel + 1
	end if
end if
'IF UCASE(rs6("DEPT")) = "FOREL" AND rs6("TYPE") = "SP" then
'totalspForel = totalspForel + 1
'end if


 ' ----------------------------------------------------------------------------------Willain Line - of glass Scanned  - Added March 10th 2014
IF UCASE(rs6("DEPT")) = "WILLIAN" then
	IF rs6("TYPE") = "SU" OR rs6("TYPE") = "OV"  then
		totalsuWillian = totalsuWillian + 1
	else 
		totalOtherWillian = totalOtherWillian + 1
	end if
end if

'IF UCASE(rs6("DEPT")) = "WILLIAN" AND rs6("TYPE") = "SP" then
'totalspWillian = totalspWillian + 1
'end if
 
  rs6.movenext
loop

rs6.close
set rs6 = nothing

totalForel = totalsuForel + totalOtherForel
totalWillian = totalsuWillian + totalOtherWillian
Response.write "GlassLine Today <br/>"
' ------------------------------------------------------------------------------------------------Awning Today

totalOVA = 0
totalOVB = 0
totalOVC = 0
totalOVD = 0
totalOVAy = 0
totalOVBy = 0
totalOVCy = 0
totalOVDy = 0

Set rs11 = Server.CreateObject("adodb.recordset")
strSQL11 = "SELECT * From X_BARCODEOV WHERE (DAY = " & cDAY & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear & ") OR  (DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary & ") ORDER BY DATETIME DESC"
rs11.Cursortype = 2
rs11.Locktype = 3
rs11.Open strSQL11, DBConnection

rs11.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs11.eof
	
	IF rs11("DEPT") = "FrameAssemble" then
		totalOVA = totalOVA + 1
	end if
	IF rs11("DEPT") = "SashAssemble" then
		totalOVB = totalOVB + 1
	end if
	IF rs11("DEPT") = "SashGlaze" then
		totalOVC = totalOVC + 1
	end if
	IF rs11("DEPT") = "WindowMount" then
		totalOVD = totalOVD + 1
	end if
rs11.movenext
loop

rs11.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary
Do while not rs11.eof
	
	IF rs11("DEPT") = "FrameAssemble" then
		totalOVAy = totalOVAy + 1
	end if
	IF rs11("DEPT") = "SashAssemble" then
		totalOVBy = totalOVBy + 1
	end if
	IF rs11("DEPT") = "SashGlaze" then
		totalOVCy = totalOVCy + 1
	end if
	IF rs11("DEPT") = "WindowMount" then
		totalOVDy = totalOVDy + 1
	end if
rs11.movenext
loop


rs11.close
set rs11 = nothing


Response.write "Awning Today <br/>"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Glazing by Opening (New April 2016)
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Set rs15 = Server.CreateObject("adodb.recordset")
strSQL15 = "Select * FROM X_GLAZING WHERE (DAY = " & cDAY & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear & ") OR  (DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary & ") Order by BARCODE ASC, FIRSTCOMPLETE DESC"
rs15.Cursortype = 2
rs15.Locktype = 3
rs15.Open strSQL15, DBConnection
rs15.filter = "DAY = " & cDay & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear


ScanCompleteWindow = 0
ScanWindow = 0

Do while not rs15.eof
	OldBarcode = Barcode
	Barcode = rs15("Barcode")
	AllScans = AllScans + 1	
	if OldBarcode = Barcode then
		if rs15("FirstComplete") = "TRUE" then
			ScanCompleteWindow = ScanCompleteWindow + 1
			ScanWindow = ScanWindow - 1
		end if
	else
		if rs15("FirstComplete") = "TRUE" then
			' ------------------------------Added Code May 3, 2016 to collect Square footage based on Job Table -------------------
			JobName = rs15("job")
			FloorName = rs15("floor")
			TagName = rs15("Tag")

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
			
			TotalSquareFoot = TotalSquareFoot + SquareFoot
			TotalWindowPerimeter = TotalWindowPerimeter + WindowPerimeter
			'--------------------------- End new Calculation ---------------------Code added at End for adding to Database ----------
			
			
			ScanCompleteWindow = ScanCompleteWindow + 1
		else
			ScanWindow = ScanWindow + 1
		end if
	end if
rs15.movenext
loop

' Now Yesterday Values
rs15.filter = ""
rs15.filter = "DAY = " & cYesterday & " AND MONTH = " & cMONTHy & " AND YEAR = " & cYeary


ScanCompleteWindowy = 0
ScanWindowy = 0

Do while not rs15.eof
	OldBarcode = Barcode
	Barcode = rs15("Barcode")
	if OldBarcode = Barcode then
		if rs15("FirstComplete") = "TRUE" then
			ScanCompleteWindowy = ScanCompleteWindowy + 1
			ScanWindowy = ScanWindowy - 1
		end if
	else
		if rs15("FirstComplete") = "TRUE" then
			' ------------------------------Added Code May 3, 2016 to collect Square footage based on Job Table -------------------
			JobName = rs15("job")
			FloorName = rs15("floor")
			TagName = rs15("Tag")

			Set rsSQFT = Server.CreateObject("adodb.recordset")
			strSQL = "Select X,Y FROM " & JobName & " Where JOB = '" &  JobName & "' and  Floor = '" &  FloorName & "' and Tag = '" &  TagName & "'"
			rsSQFT.Cursortype = 2 
			rsSQFT.Locktype = 3
			On Error Resume Next  
			rsSQFT.Open strSQL, DBConnection
			If rsSQFT.State = 1 Then 
				SquareFoot = rsSQFT("X") * rsSQFT("Y") / 144
				WindowPerimeter = (rsSQFT("X") * 2 ) + (rsSQFT("Y") * 2)
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
		else
			ScanWindowy = ScanWindowy + 1
		end if
	end if
rs15.movenext
loop

rs15.close
set rs15 = nothing

Response.write "Glazing Today <br/>"

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Second Set set of Todays Data Complete - Total Job Values + Today's Work
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------


' -----------------------------------------------------------------------------------------------------------------FILL IN X_BARCODE LINE ITEM



Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_BARCODE_LINEITEM WHERE DEPT ='ASSEMBLY' ORDER BY JOB, FLOOR ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection	

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection


DIM JFCHECKID, wcount
'Declared Variables for the Today's Activities section
wcount=0
' Reset the filter on X_Barcode back to full (in the previous section, it was filtered to yesterday)
rs.filter = ""

'THIS IS SLOW, BUT IT MAKES SURE THAT ALL WINDOWS RELATED TO A BATCH ARE COUNTED FROM MORE THEN TODAY'S LIST
if not rs.eof then
rs.movefirst 
end if
do while not rs.eof
'Filter out RS2 items that do not have an ID
rs2.filter = "ID > 0 "
' Count =1 Means already counted from Barcode, so skip,  Any other value is to be counted here and then COUNT SET to 1 
	IF RS("COUNT") <> 1 Then 

		JFCHECKID = 0
		rs2.filter =" JOB = '" &  UCASE(rs("JOB")) & "' AND FLOOR = '" &  UCASE(rs("FLOOR")) & "' AND DEPT = '" &  UCASE(rs("DEPT")) & "'"
		if not rs2.eof then
				JFCHECKID = RS2("ID")
				wcount = wcount + 1	
		END IF			
	
			
		
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
							
							
							
							rs2.fields("YEAR") = RS("YEAR")
							rs2.fields("MONTH") = RS("MONTH")
							rs2.fields("DAY") = RS("DAY")
							rs2.fields("WEEK") = RS("WEEK")
							
							rs2.update
						end if 
					end if
					
				
				wcount = 0
	' Flags the Count in Barcode to 1 so that it is not counted again			
	RS("COUNT") = 1
	End If
	rs.movenext
loop	


' -------------------------------------------------------------------------------------------------------- Today's Values from X_BARCODE_LINEITEM
TodayJobStat = ""
strSQL14 = "DELETE From V_REPORT2 WHERE DEPT = 'ASSEMBLY'"
Set rs14 = DBConnection.Execute(strSQL14)

' Today Count for each job
rs.filter =""
rs2.filter=""
rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
rs2.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
if not rs2.eof then
rs2.movefirst 
end if
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
	rs2.update
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
			rs4.update	
		end if
		if NOT rs2("DEPT") = "GLAZING2" then
			'  Jody Reuqested that Glazing2 be removed from this section June 2014
			' Not filtered above so that X_Barcode_lineitem still collects the data, only the showing is removed.
			
			TodayJobStat = TodayJobStat &  "<li>" & rs2("DEPT") & " " & UCASE(rs2("Job")) & " " & rs2("Floor") & ": (" & rs2("Last") & ") " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
			
			
			'CODE TO WRITE IT TO THE RS13 TABLE WHICH IS CALLED V_REPORT2 TO WRITE ACTIVITY TO THE DATABASE



strSQL13 = "INSERT INTO V_REPORT2 (DEPT, JOB, FLOOR, TODAY, PROGRESS, TOTAL) VALUES ('" & RS2("DEPT") & "', '" & UCASE(rs2("Job")) & "', '" & RS2("FLOOR") & "', '" & RS2("LAST") & "', '" & RS2("TAG") & "','" &  RS4("TOTALWIN") & "')"
Set rs13 = DBConnection.Execute(strSQL13)

		
		end if

	end if
	rs2.movenext
	
	loop

	
rs.close
set rs=nothing
rs2.close
set rs2=nothing

rs4.close
set rs4 = nothing



'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' THIRD Set set of Todays Data  - MACHINE STATUS
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------


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
strSQL10 = "SELECT * From PROECOVERT2 ORDER BY ID DESC"
rs10.Cursortype = 2
rs10.Locktype = 3
rs10.Open strSQL10, DBConnection

						EWWIDTH =  rs7("jobNumber") & " " & rs7("Cutstatus") & "%"
						EWJAMB = rs8("jobNumber") & " " & rs8("Cutstatus") & "%"
						QWIDTH = rs9("jobNumber") & " " & rs9("Cutstatus") & "%"
						QJAMB = rS10("jobNumber") & " " & rs10("Cutstatus") & "%"
rs7.close
set rs7=nothing
rs8.close
set rs8=nothing
rs9.close
set rs9=nothing
rs10.close
set rs10=nothing					



Response.write "Machine Status <br/>"

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fourth Set - ZIPPER MACHINE CUTS
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

Response.write "Zippers<br/>"

'--------------------------------------------------------------------------------------------------------------------------------------
' Move X_BARCODE TO X_WIN PROD
'------------------------------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Glazing by Opening Full and Partial (BY Job for Activity Values
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------


Set rs1 = Server.CreateObject("adodb.recordset")
strSQL1 = "Select * FROM X_GLAZING Where FIRSTCOMPLETE = 'TRUE' order by JOB,FLoor ASC"
rs1.Cursortype = 2
rs1.Locktype = 3
rs1.Open strSQL1, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "Select * FROM X_BARCODE_LINEITEM WHERE DEPT ='GLAZING' ORDER BY ID DESC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection	

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

	rs1.filter=""
		
	Do WHile not rs1.eof
		IF RS1("COUNT") <> 1 Then 
		JFCHECKID = 0
		rs3.filter = ""
		rs3.filter =" JOB = '" &  UCASE(rs1("JOB")) & "' AND FLOOR = '" &  UCASE(rs1("FLOOR")) & "' AND DEPT = '" &  UCASE(rs1("DEPT")) & "'"
		if not rs3.eof then
				JFCHECKID = RS3("ID")
				wcount = wcount + 1	
		END IF
			
		IF JFCHECKID = "0" THEN
			rs3.addnew 
			rs3.fields("JOB") = UCASE(RS1("JOB"))
			rs3.fields("FLOOR") = RS1("FLOOR")
			rs3.fields("DEPT") = RS1("DEPT")
			rs3.fields("YEAR") = RS1("YEAR")
			rs3.fields("MONTH") = RS1("MONTH")
			rs3.fields("DAY") = RS1("DAY")
			rs3.fields("WEEK") = RS1("WEEK")
			rs3.fields("TAG") = 1
					
			rs3.UPDATE
		ELSE	
					
			if JFCHECKID = "" then
				response.write "<li>GOTCHA NULL</li>"
			else
				rs3.filter = "ID = " & JFCHECKID
				rs3.fields("TAG") = rs3.Fields("TAG") + wcount
							
				rs3.fields("YEAR") = RS1("YEAR")
				rs3.fields("MONTH") = RS1("MONTH")
				rs3.fields("DAY") = RS1("DAY")
				rs3.fields("WEEK") = RS1("WEEK")
				rs3.update							
				
			end if 
		end if
		Add = ADD + 1
	wcount = 0
	' Flags the Count in Barcode to 1 so that it is not counted again			
	RS1("COUNT") = 1
	End If
rs1.movenext
loop

'--------------------------------------------------------------------------------------------------------------------------------------
' Move X_BARCODE TO X_WIN PROD
'------------------------------------------------------------------------------------------------------------------------------------

rs3.filter = "MONTH = '" & cMONTH & "' AND YEAR = '" & cYear & "'"
Do While not rs3.eof
			rs4.filter =""
			rs4.filter = "JOB = '" & UCASE(RS3("JOB")) & "' AND FLOOR = '" & RS3("FLOOR") & "'"
			if not rs4.eof then
				rs4("DONEWING") = rs3("TAG")
				rs4.update
			end if

rs3.movenext
loop

'------------------------------------------------------------------------------------------------------------------------------------
'TODAY Activity report for V_REPORT2
'------------------------------------------------------------------------------------------------------------------------------------


TodayJobStat = ""
strSQLV2D = "DELETE From V_REPORT2 WHERE DEPT = 'GLAZING'"
Set rsV2D = DBConnection.Execute(strSQLV2D)

' Today Count for each job
rs1.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
rs3.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
if not rs3.eof then
rs3.movefirst 
end if
do while not rs3.eof
'Filter out RS3 items that do not have an ID

	todayjob = 0
	
	rs1.movefirst 
	do while not rs1.eof
		if rs1("JOB") = rs3("JOB") AND rs1("Floor") = rs3("FLOOR") AND rs1("DEPT") = rs3("DEPT") then
		todayjob = todayjob + 1
		end if
	rs1.movenext
	loop
	rs3("LAST") = todayjob
	rs3.update
rs3.movenext
loop
rs1.filter = ""
rs3.filter = ""

	'Filters all the information that is in X_BARCODE_LINEITEM to just show items worked on Today
	rs3.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear 
	do while not rs3.eof
	rs4.filter = "Job = '" &  UCASE(rs3("Job")) & "' AND Floor = '" &  rs3("Floor") & "'"
	if rs4.bof then
	else
		IF rs3("DEPT") = "GLAZING" then
			rs4.fields("DoneWinG") = rs3.fields("Tag")
			rs4.update	
		end if

			TodayJobStat = TodayJobStat &  "<li>" & rs3("DEPT") & " " & UCASE(rs3("Job")) & " " & rs3("Floor") & ": (" & rs3("Last") & ") " & rs3("Tag") & "/" & rs4("TotalWin") & "</li>"
			
			
			'CODE TO WRITE IT TO THE RS13 TABLE WHICH IS CALLED V_REPORT2 TO WRITE ACTIVITY TO THE DATABASE



strSQLV2 = "INSERT INTO V_REPORT2 (DEPT, JOB, FLOOR, TODAY, PROGRESS, TOTAL) VALUES ('" & RS3("DEPT") & "', '" & UCASE(rs3("Job")) & "', '" & RS3("FLOOR") & "', '" & RS3("LAST") & "', '" & RS3("TAG") & "','" &  RS4("TOTALWIN") & "')"
Set rsV2 = DBConnection.Execute(strSQLV2)

		
	

	end if
	rs3.movenext
	
	loop

	
rs1.close
rs3.close
set rs1 = nothing
set rs3 = nothing

rs4.close
set rs4 = nothing


'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fifth Set SHIPPING DATA 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Set rs16 = Server.CreateObject("adodb.recordset")
strSQL16 = "Select * FROM X_SHIP"
rs16.Cursortype = 2
rs16.Locktype = 3
rs16.Open strSQL16, DBConnection	
rs16.filter = "ShipDate = #" & CurrentDate & "#"

ShipScan = 0
ShipScany = 0
Do while not rs16.eof
	ShipScan = ShipScan + 1
rs16.movenext
loop

rs16.filter =""
rs16.filter = "ShipDate = #" & YesterdayDate & "#"

Do while not rs16.eof
	ShipScany = ShipScany + 1
rs16.movenext
loop



Set rs17 = Server.CreateObject("adodb.recordset")
strSQL17 = "Select top 100 * FROM X_SHIP_TRUCK order by id DESC"
rs17.Cursortype = 2
rs17.Locktype = 3
rs17.Open strSQL17, DBConnection

TruckOpen = 0
TruckClose = 0
TruckOpenName = ""
TruckCloseName = ""

TruckOpeny = 0
TruckClosey = 0
TruckOpenNamey = ""
TruckCloseNamey = ""



Do while not rs17.eof
	
	if rs17("Shipdate") = currentDate then 
		TruckClose = TruckClose + 1
		TruckCloseName = TruckCloseName & rs17("sList") & " | "
	end if
	
	if rs17("Createdate") = currentDate then 
		TruckOpen = TruckOpen + 1
		TruckOpenName = TruckOpenName & rs17("sList") &  " | "
	end if
	
	if rs17("Shipdate") = yesterdayDate then 
		TruckClosey = TruckClosey + 1
		TruckCloseNamey = TruckCloseNamey & rs17("sList") &  " | "
	end if
	
	if rs17("Createdate") = yesterdayDate then 
		TruckOpeny = TruckOpeny + 1
		TruckOpenNamey = TruckOpenNamey & rs17("sList") &  " | "
	end if
	
rs17.movenext
loop




rs16.close
rs17.close
set rs16 = nothing
set rs17 = nothing


Response.write "Shipping Today<br/>"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Sixth Set INSERT DATA 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------



Set rs12 = Server.CreateObject("adodb.recordset")
strSQL12 = "SELECT * From V_REPORT1"
rs12.Cursortype = 2
rs12.Locktype = 3
rs12.Open strSQL12, DBConnection


rs12.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear

if rs12.eof then
	rs12.addnew
end if
rs12.fields("GLAZING") = totalg
rs12.fields("SquareFoot") = TotalSquareFoot
rs12.fields("WindowPerimeter") = TotalWindowPerimeter
rs12.fields("GLAZING2") = totalg2
rs12.fields("ASSEMBLY") = totala
rs12.fields("PANEL") = totalP
rs12.fields("AWNING") = totalOVC
rs12.fields("FOREL") = totalForel
rs12.fields("WILLIAN") = totalWillian
rs12.fields("TIMEFRAME") = "TODAY"

rs12.fields("EWWIDTH") = EWWIDTH 
rs12.fields("EWJAMB") = EWJAMB 
rs12.fields("QWIDTH") = QWIDTH 
rs12.fields("QJAMB") = QJAMB
rs12.fields("DAY") = cDay
rs12.fields("MONTH") = cMonth
rs12.fields("YEAR") = cYear

rs12.fields("ZIPPERRED") = ZipperRed
rs12.fields("ZIPPERBLUE") = ZipperBlue

rs12.fields("ShipScan") = ShipScan
rs12.fields("TruckOpen") = TruckOpen
rs12.fields("TruckClose") = TruckClose
rs12.fields("TruckOpenName") = TruckOpenName
rs12.fields("TruckCloseName") = TruckCloseName

rs12("GlazingFull") = ScanCompleteWindow
rs12("GlazingPartial") = ScanWindow

rs12.UPDATE

rs12.filter = ""
rs12.filter = "DAY = " & cYesterday & " AND MONTH = " & cMonthy & " AND YEAR = " & cYeary

if rs12.eof then
	rs12.addnew
	rs12.fields("DAY") = cYesterday
	rs12.fields("MONTH") = cMonthy
	rs12.fields("YEAR") = cYeary
	rs12.fields("TIMEFRAME") = "YESTERDAY"
end if

rs12.fields("GLAZING") = totalgy
rs12.fields("SquareFoot") = TotalSquareFooty
rs12.fields("WindowPerimeter") = TotalWindowPerimetery
rs12.fields("GLAZING2") = totalg2y
rs12.fields("ASSEMBLY") = totalay
rs12.fields("AWNING") = totalOVCy
rs12.fields("PANEL") = totalPy
rs12.fields("ZIPPERRED") = ZipperRedy
rs12.fields("ZIPPERBLUE") = ZipperBluey

rs12.fields("ShipScan") = ShipScany
rs12.fields("TruckOpen") = TruckOpeny
rs12.fields("TruckClose") = TruckClosey
rs12.fields("TruckOpenName") = TruckOpenNamey
rs12.fields("TruckCloseName") = TruckCloseNamey

rs12("GlazingFull") = ScanCompleteWindowy
rs12("GlazingPartial") = ScanWindowy

rs12.update

rs12.close
set rs12=nothing

DBConnection.close
set DBConnection=nothing

%>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<b><u>Glass Production Stats</u></b>
		<br>
		<li><% response.write "FULL GLAZING: " & ScanCompleteWindow & " - " & TotalSquareFoot & "ft<sup>2</sup>"%></li>
		<li><% response.write "PARTIAL GLAZING: " & ScanWindow %></li>
<!--
		<li><% response.write "GLAZING: " & totalg %></li>
		<li><% response.write "Yesterday GLAZING: " & totalgy %></li>
		<li><% response.write "GLAZING2: " & totalg2 %> </li>
-->
		<li><% response.write "ASSEMBLY: " & totala %></li>
		<li><% response.write "PANEL: " & totalP %></li>
		<li><% response.write "AWNING: " & totalOVC %></li>

		<br>
		<b><u>Glassline Stats</u></b>
		<li><% response.write "FOREL: " & totalForel %></li>
		<li><% response.write "WILLIAN: " & totalWillian %></li>
		<br>

		<br>
		<b><u>Zipper Stats</u></b>
		<li><% response.write "Zipper Red: " & ZipperRed %></li>
		<li><% response.write "Zipper Blue: " & ZipperBlue %></li>

		<br>
		<b><u>Shipping Stats</u></b>
		<li><% response.write "Scan (Today/Yesterdat): " & ShipScan & " / " & ShipScany %></li>
		<li><% response.write "Open Truck (Today/Yesterday): " & TruckOpen & " / " & TruckOpeny %></li>
		<li><% response.write "Open Truck (Today/Yesterday): " & TruckOpenName & " / " & TruckOpenNamey %></li>
		<li><% response.write "Close Truck (Today/Yesterday): " & TruckClose & " / " & TruckClosey %></li>
		<li><% response.write "Close Truck (Today/Yesterday): " & TruckCloseName & " / " & TruckCloseNamey %></li>

		<br>
		<b><u>Awning Stats - In Development</u></b>
		<li><% response.write "Frame Assembly: " & totalOVA %></li>
		<li><% response.write "Sash Assembly: " & totalOVB %></li>
		<li><% response.write "Sash Glazing: " & totalOVC %></li>
		<li><% response.write "Window Mounting: " & totalOVC %></li>
		<br>
       <b><u>Today's Activity</u></b>

		<% response.write TodayJobStat %>

		<br>
		<b><u>CNC Machines</u></b>
<%
	RESPONSE.WRITE "<li>EW WIDTH: " & EWWIDTH & "</li>"
	RESPONSE.WRITE "<li>Q WIDTH: " & QWIDTH & "</li>"
RESPONSE.WRITE "<li>EW Vert: " & EWJAMB & "</li>"
RESPONSE.WRITE "<li>EW Vert2: " & QJAMB & "</li>"
%>

        </ul>

</body>
</html>
