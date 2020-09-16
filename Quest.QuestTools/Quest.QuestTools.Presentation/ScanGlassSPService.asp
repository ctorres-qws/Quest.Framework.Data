<!--#include file="dbpath.asp"-->              
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
           
			<!-- SCAN to complete Glass - from Backorder and from Glassline for Manual SP Service Glass-->
			<!-- Created May 2015 as duplicate of Scan Glass Forel -->
			
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Scan</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <meta http-equiv="refresh" content="1000" >
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  </script>
 
    
     <!--#include file="TodayAndYesterday.asp"-->
<%
Commacount=0

error = ""

DEPTVAR = "SP-SERVICE"
glassScanNum = 0
bc = UCASE(request.querystring("window"))
glassScanNum = request.querystring("GSN")
if glassScanNum = "" then
	glassScanNum = 0
end if


OUTDATE = DATE
'The date to add on the item marked completed
IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found on the Z_GLASSDB
DetailsID =""
' Details to show for item successfully marked complete
RecordLocated=0
' Add to determine if UNSCANED record already Exists

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

' Edit April 2015 - Scan within 5 days (Faster search and allows recuts of older jobs without change)
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("Select * FROM X_BARCODEGA WHERE DATETIME > #" & DATE()-5 & "# ", isSQLServer)
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Exception Scans GT (Backorder) AND * (Service)
if Left(bc,2) = "GT" AND NOT Left(bc,3) ="GTM" OR Left(bc,1) = "."then

' GT code means Backorder, so the code will run first, if the barcode is not a backorder, go down to the else
' There is a patch in place, GT is code for BackOrder, so GTM had to be hardcoded or it ran like a backorder
	if Left(bc,2) = "GT" AND NOT Left(bc,3) ="GTM" then
		if len(bc) <3 then 
			bc = "GT00"
		end if

	'Drop the GT in front of the ID to get the ID number
		GlassID = Mid(bc, 3)

		Set rs4 = Server.CreateObject("adodb.recordset")
		strSQL4 = "SELECT top 1000 * FROM Z_GLASSDB WHERE ID=" & GlassID & " ORDER BY ID ASC"
		rs4.Cursortype = 2
		rs4.Locktype = 3
		rs4.Open strSQL4, DBConnection

		Do while not rs4.eof
			if CLng(rs4("ID")) = CLng(GlassID) then
				if isDate(rs4.fields("CompletedDate")) = False then
					' Record in Z_GLASSDB matches the scanned item and does not have an Output Date
					' Successfully Marked Done
					rs4.fields("CompletedDate") = OUTDATE
					rs4.update
					' Details of Completed Item to be shown
					JOB = rs4.fields("JOB")
					FLOOR = rs4.fields("FLOOR")
					TAG = rs4.fields("TAG")
					DetailsID = "JOB: " &JOB & " FLOOR: " & FLOOR & " TAG: " & TAG
				
					'Declare additional Variable logic
					' LOGIC to Determine GLASSTYPE
						INTERIORW = rs4.fields("1 MAT")
						EXTERIORW = rs4.fields("2 MAT")
					
						if INTERIORW = "-" then 
							GLASSTYPE = "SP"
						else 
							GLASSTYPE = "SU"
						end if
				
				'CHeck to see if the record already exists in X_BARCODEGA and has UNSCAN as Department
				' If it does, then just change the Department back to DEPTVAR
				' If it does not, then add a new record
				' Record Located = ID
				
					rs.filter = "BARCODE = '" & bc & "' AND DEPT = 'UNSCAN'"
					if not rs.eof then 
						RecordLocated = rs("ID")
					end if
					rs.filter = ""
				
					rs.movefirst
					If RecordLocated > 0 then
						' If the Record was found with an UNSCAN then filter to that ID
						' If the Record was not located then add a new one in the else
				
						rs.Find "ID = " & RecordLocated
					else
			
						'Add a Record to X_BARCODEGA code
						'rs = "Select * FROM X_BARCODEGA" from include

						rs.addnew 
					end if
				
					rs.fields("BARCODE") = bc
					rs.fields("JOB") = JOB
					rs.fields("FLOOR") = FLOOR
					rs.fields("TAG") = TAG
					rs.fields("DEPT") = DEPTVAR
					rs.fields("DATETIME") = STAMPVAR
					rs.fields("TYPE") = GLASSTYPE
					rs.fields("TIME") = cctime
				
					if hour(now) <= 6 then  ' Changed to 6am from 3 by Michael Bernholtz February 2018
						rs.fields("DAY") = cYesterday
						rs.fields("MONTH") = cmonthy
						rs.fields("YEAR") = cyeary
						rs.fields("WEEK") = weeknumbery
					else
						rs.fields("DAY") = cday
						rs.fields("MONTH") = cmonth
						rs.fields("YEAR") = cyear
						rs.fields("WEEK") = weeknumber
					end if
					
				
					rs.UPDATE
				
				else 	
					' Record in Z_GLASSDB matches the scanned item  but already has an Output Date
					' Previously Marked Done
					error = "Glass Order: " & GlassID & " already marked completed." 
					IsError = True
				end if 
				IDFound = True	
			end if
		rs4.movenext
		loop
		rs4.close
		set rs4=nothing
		if IDFound = False then 
			' Record not found in Z_GLASSDB for comparison
			error =" Scanned ID: " & GlassID & " does not exist."
			IsError = True
		end if

	else
	' * Means Service Barcode - Service Barcode has .TypePO.POLine
		
		
			
	sizecheckid = 0
	
	' Attempt to find existing record and cancel if exist or rescan if marked Unscan
	' This code updated based on filter instead of rs loop
	rs.filter = "BARCODE = '" & bc & "' AND DEPT = '" & DEPTVAR & "'"
		if not rs.eof then 
			sizecheckid = rs("ID")
			error = bc & ": Already Scanned - Not Sent"
			IsError = True
		end if
	rs.filter = ""
	rs.filter = "BARCODE = '" & bc & "' AND DEPT = 'UNSCAN'"
	if not rs.eof then 
			RecordLocated = rs("ID")
		end if
	rs.filter = ""
		  
		if sizecheckid = 0 then

			if Len(bc) > 3 then

			GLASSTYPE = RIGHT(LEFT(bc,3),2)
			POLINE = Mid(bc, instr(3,bc, ".")+1,5)
			PO = Mid(bc, 4, instr(3,bc, ".")-4)
			
				rs.movefirst
				If RecordLocated > 0 then
					' If the Record was found with an UNSCAN then filter to that ID
					' If the Record was not located then add a new one in the else
				
					rs.Find "ID = " & RecordLocated
				else
					'Add a Record to X_BARCODEGA code
					'rs = "Select * FROM X_BARCODEGA" from include
					rs.addnew 
				end if
				
				rs.fields("BARCODE") = bc
				rs.fields("DEPT") = DEPTVAR
				rs.fields("DATETIME") = STAMPVAR
				rs.fields("TYPE") = GLASSTYPE
				rs.fields("PO") = PO
				rs.fields("POLINE") = POLINE
				rs.fields("TIME") = cctime
				
				if hour(now) <= 3 then
					rs.fields("DAY") = cYesterday
					rs.fields("MONTH") = cmonthy
					rs.fields("YEAR") = cyeary
					rs.fields("WEEK") = weeknumbery
				else
					rs.fields("DAY") = cday
					rs.fields("MONTH") = cmonth
					rs.fields("YEAR") = cyear
					rs.fields("WEEK") = weeknumber
				end if
				
			
				rs.UPDATE
			
			else 
				error = bc & ": Not a Valid Barcode, Try Again"
				IsError = True
			end if

		end if
	end if
' Normal Read Barcode after Exceptions 
else
' *************************************************************************
	Endlineg = "0"
	if LEN(BC) >1 then
		Shortbc = left(bc, LEN(bc)-1)
	end if
	' Error arose when MU created 2 "1SU" so Position and Glass type can now hold 4 characters 1SU1, 1SU2
	if Right(bc, 2) = "OV" OR Right(bc, 2) = "SU" OR Right(bc, 2) = "SP" OR Right(bc, 2) = "TG" OR Right(bc, 2) = "OV" OR Right(bc, 2) = "HS" OR Right(bc, 2) = "SB" OR Right(shortbc, 2) = "OV" OR Right(shortbc, 2) = "SU" OR Right(shortbc, 2) = "SP" OR Right(shortbc, 2) = "TG" OR Right(shortbc, 2) = "OV" OR Right(shortbc, 2) = "HS" OR Right(shortbc, 2) = "SB" OR Right(shortbc, 2) = "SW" then
		if Right(bc, 2) = "OV" OR Right(bc, 2) = "SU" OR Right(bc, 2) = "SP" OR Right(bc, 2) = "TG" OR Right(bc, 2) = "OV" OR Right(bc, 2) = "HS" OR Right(bc, 2) = "SB" OR Right(bc, 2) = "SW" then
			glasstype = right(bc,2)
			Endlineg = Right(bc, 3)
			Position = Left(Endlineg,1)
			bc = left(bc, LEN(bc)-3)
		else 
			glasstype = left(right(bc,3),2)
			Endlineg = Right(bc, 4)
			Position = Left(Endlineg,1)
			bc = left(bc, LEN(bc)-4)
		end if 	
	else
		glasstype = "00"
		Position = "0"
	end if
	
	jobname = Left(bc, 3)
	if inStr(1, bc, "-", 0) = 5 then
		floor = Mid(bc, 4, 1)
		tag = Mid(bc, 6, 7)
	END IF

	if inStr(1, bc, "-", 0) = 6 then
		floor = Mid(bc, 4, 2)
		tag = Mid(bc, 7, 7)
	end if

	if inStr(1, bc, "-", 0) = 7 then
		floor = Mid(bc, 4, 3)
		tag = Mid(bc, 8, 7)
	end if
	
	if inStr(1, bc, "-", 0) = 8 then
		floor = Mid(bc, 4, 4)
		tag = Mid(bc, 9, 7)
	end if
	
	if inStr(1, bc, "-", 0) = 9 then
		floor = Mid(bc, 5, 5)
		tag = Mid(bc, 10, 7)
	end if

	if Endlineg = "0" then
	else
		bc = bc & Endlineg
	end if

	sizecheckid = 0

	' Attempt to find existing record and cancel if exist or rescan if marked Unscan
	' This code updated based on filter instead of rs loop, March 23, 2015
	rs.filter = "BARCODE = '" & bc & "' AND DEPT = '" & DEPTVAR & "'"
		if not rs.eof then 
			sizecheckid = rs("ID")
			error = bc & ": Already Scanned - Not Sent"
			IsError = True
		end if
	rs.filter = ""
	rs.filter = "BARCODE = '" & bc & "' AND DEPT = 'UNSCAN'"
	if not rs.eof then 
			sizecheckid = rs("ID")
			RecordLocated = rs("ID")
			rs.Move RecordLocated
			rs.fields("DEPT") = DEPTVAR
			rs.UPDATE
			DetailsID = "JOB: " &jobname & " FLOOR: " & floor & " TAG: " & tag
		end if
	rs.filter = ""
	BarcodeSplit = split(BC,"-")
	CommaCount = Ubound(BarcodeSplit)
	if CommaCount>3 then
		jobname = BarcodeSplit(0)
		floor = BarcodeSplit(1)
		tag = BarcodeSplit(2)
		Position = BarcodeSplit(3)
	end if
	
	if sizecheckid = 0 then

		if Len(bc) > 3 then

			rs.addnew 
			rs.fields("BARCODE") = bc
			rs.fields("JOB") = jobname
			rs.fields("FLOOR") = floor
			rs.fields("TAG") = tag
			rs.fields("POSITION") = Position
			rs.fields("DEPT") = DEPTVAR
			rs.fields("DATETIME") = STAMPVAR
			rs.fields("TYPE") = glasstype
			rs.fields("TIME") = cctime
			'Any time from 2 am and before is counted as previous day
			if hour(now) <= 3 then
				rs.fields("DAY") = cYesterday
				rs.fields("MONTH") = cmonthy
				rs.fields("YEAR") = cyeary
				rs.fields("WEEK") = weeknumbery
			else
				rs.fields("DAY") = cday
				rs.fields("MONTH") = cmonth
				rs.fields("YEAR") = cyear
				rs.fields("WEEK") = weeknumber
			end if

			rs.UPDATE

			DetailsID = "JOB: " &jobname & " FLOOR: " & floor & " TAG: " & tag

		else 
			error = bc & ": Not a Valid Barcode, Try Again"
			IsError = True
		end if

	end if
end if

rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing

DbCloseAll

End Function

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">SP SCAN</h1>
        <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="SP Scan" class="panel" name="igline" action="ScanGlassSPService.asp" method="GET" selected="true">
         <h1><center>SP Glass Scan</center></h1>
		 <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <p>
				<% if IsError = False then
				response.write "<img src='complete.jpg' alt='Scan Completed - Thank You' width='70' height='70'><BR>"
				response.write bc & " - Sent"
				glassScanNum = glassScanNum + 1
				response.write " --> Window # " & glassScanNum 
				response.write "<BR><img src='complete.jpg' alt='Scan Completed - Thank You' width='75' height='75'>"

				else
				response.write error
					Response.write ":" & CommaCount & ":"
				end if				
				%></p>
              
            </div>
            <% 			
			end if %>
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
            <input type="hidden" name='GSN' id='GSN' value="<%response.write glassScanNum %>" />
                              	

            
            
                  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
    
        textbox.value = barcode;
		igline.submit();
    
}
			
			

    

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
            
            
            
            </form>
			
<!--Ending rs,rs4, DBConnection at last location -->

</body>
</html>