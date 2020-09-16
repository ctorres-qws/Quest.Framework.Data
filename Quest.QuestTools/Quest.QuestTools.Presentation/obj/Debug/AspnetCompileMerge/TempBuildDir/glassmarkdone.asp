<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 5th, 2013 - by Michael Bernholtz at request of Jody Cash --> 

		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Mark Done</title>
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
strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'afilter = request.QueryString("aisle")

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
        <ul id="Profiles" title="Mark Glass Order Done" selected="true">

<%

	MARKEDID = REQUEST.QueryString("ID")
	UNDO = REQUEST.QueryString("UNDO")
	OUTDATE = Date

'rs.filter = "WAREHOUSE='GOREWAY'"
'rs.filter = "AISLE='" & afilter & "'"

'Undo button only appears after an item is marked done
if CLng(UNDO) <> 0 then
	response.write "<li class='group'><a href='glassmarkdone.asp?ID=" & MARKEDID & "&UNDO=0' target='_self'> UNDO - " & MARKEDID & "</a></li>"
end if
response.write "<li>Select a Glass Item to Mark Done:</li>"

do while not rs.eof
	MainId = rs.fields("ID")
'Declare BARCODE so it can be used in both the Entry and the UNDO
	WBARCODE = rs.Fields("BARCODE")

' Marked item is set to Mark as Done (and add an Output Date) if Undo is 1
' Marked item is set to Unmark as Done (and return Output Date to Null) if Undo is 0
	if CLng(UNDO) = 1 then

		if CLng(MainId) = CLng(MARKEDID) then 
			' Marks Complete
			rs.Fields("COMPLETEDDATE") = OUTDATE
	
			'Add a Record to X_BARCODEGA similar to the GA1.asp code
				'Declare Variables for Job Floor Tag based on rs.
					JOB = rs.fields("JOB")
					FLOOR = rs.fields("FLOOR")
					TAG = rs.fields("TAG")
				'Declare variables to add new record at current Date
					STAMPVAR = year(now) & " " & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now)
					ccTime = hour(now) & ":" & minute(now)
					cDay = day(now)
					cMonth = month(now)
					cYear = year(now)
					currentDate = Date
					weekNumber = DatePart("ww", currentDate)
				' Declare variable for Department 
					DEPTVAR = "GLASSLINE"
				' LOGIC to Determine GLASSTYPE
					INTERIORW = rs.fields("1 MAT")
					EXTERIORW = rs.fields("2 MAT")

					if INTERIORW = "-" then 
						GLASSTYPE = "SP"
					else 
						GLASSTYPE = "SU"
					end if

				Set rs2 = Server.CreateObject("adodb.recordset")
				strSQL2 = "Select * FROM X_BARCODEGA"
				rs2.Cursortype = 2
				rs2.Locktype = 3
				rs2.Open strSQL2, DBConnection

				rs2.addnew 
				rs2.fields("BARCODE") = WBARCODE
				rs2.fields("JOB") = JOB
				rs2.fields("FLOOR") = FLOOR
				rs2.fields("TAG") = TAG
				rs2.fields("DEPT") = DEPTVAR
				rs2.fields("DATETIME") = STAMPVAR
				rs2.fields("TYPE") = GLASSTYPE
				rs2.fields("DAY") = cday
				rs2.fields("MONTH") = cmonth
				rs2.fields("YEAR") = cyear
				rs2.fields("WEEK") = weeknumber
				rs2.fields("TIME") = cctime
				rs2.UPDATE

				rs2.close
				set rs2 = nothing
		end if
	else
		if CLng(MainId) = CLng(MARKEDID) then 
			'Marks Uncomplete
			rs.Fields("COMPLETEDDATE") = NULL

			' Delete the Record from the X_BARCODEGA - removed
	'		Set rs2 = Server.CreateObject("adodb.recordset")
	'		strSQL2 = "DELETE * FROM X_BARCODEGA WHERE BARCODE = '" & WBARCODE & "' "
	'		rs2.Cursortype = 2
	'		rs2.Locktype = 3
	'		rs2.Open strSQL2, DBConnection

		end if
	end if
	Flag = rs.Fields("COMPLETEDDATE")
	if isNull(Flag)then
' Checks to see if Previously Marked Done: "COMPLETEDDATE" is the Output Date - It will only be filled in after selected in the display below
' Sets the ID to the ID of the item selected and moves the UNDO flag to 1 in order to allow an UNDO command to be processed
		response.write "<li><a href='glassmarkdone.asp?ID=" & MainId & "&UNDO=1' target='_self'> ID: "& rs("ID") & " " & rs("JOB") & " " & RS("FLOOR") &" " & RS("TAG") & " w " & RS("DIM X") & "'' x h " & RS("DIM Y") & "'' " & "IN " & RS("INPUTDATE") & " OPT " & RS("OPTIMADATE") & " REQ " & RS("REQUIREDDATE") & " OUT " & RS("COMPLETEDDATE") & "</a></li>"

	end if

	rs.movenext
loop

%>
      </ul>
<%
rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
%>
</body>
</html>
