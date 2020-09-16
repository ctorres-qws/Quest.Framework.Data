<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- include connect_barcodega removed-->          
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
' Declare the Variables    	 
OUTDATE = DATE
'The date to add on the item marked completed
IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found on the Z_GLASSDB
DetailsID =""
' Details to show for item successfully marked complete

'Collect the ID submitted in the form
ScannedID = request.querystring("SCANNEDID")


if Left(ScannedID,2) = "GT" then
	if len(ScannedID) <3 then 
		bc = "GT00"
	end if

	 'Drop the GT in front of the ID to get the ID number
	GlassID = Mid(ScannedID, 3)

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection


	Do while not rs.eof
		if CLng(rs("ID")) = CLng(GlassID) then
			if isDate(rs.fields("COMPLETEDDATE")) = False then
				' Record in Z_GLASSDB matches the scanned item and does not have an Output Date
				' Successfully Marked Done
				rs.fields("COMPLETEDDATE") = OUTDATE
				' Details of Completed Item to be shown
				JOB = rs.fields("JOB")
				FLOOR = rs.fields("FLOOR")
				TAG = rs.fields("TAG")
				DetailsID = "JOB: " &JOB & "FLOOR: " & FLOOR & "TAG: " & TAG
				
				'Add a Record to X_BARCODEGA similar to the GA1.asp code
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
				rs2.fields("BARCODE") = ScannedID
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

				
			else 	
				' Record in Z_GLASSDB matches the scanned item  but already has an Output Date
				' Previously Marked Done
				error = "Glass Order: " & GlassID & " already marked completed." 
				IsError = True
			end if 
			IDFound = True	
		end if
	rs.movenext
	loop
  
	if IDFound = False then 
		' Record not found in Z_GLASSDB for comparison
		error =" Scanned ID: " & GlassID & " does not exist."
		IsError = True
	end if

else
	'Scanned ITEM does not start with GT - Invalid item to scan
	error = ScannedID & ": Not Valid GT:ID."
	IsError = True
	GlassID = 0
end if  

rs.close
Set rs=nothing		
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_webapp">Glass Tools</a>
        </div>

    <form id="igline" title="Glass Line Scan" class="panel" name="igline" action="glassmarkdoneScan.asp" method="GET" selected="true">
<% 
	If ScannedID = "" Then
		 response.write ""
	Else
%>
			<div class="row">
				<label>
<%
		If IsError = False Then
			Response.write GlassID & " - Marked Done: " & DetailsID
		Else
			Response.write error
		End If
%></label>

            </div>
<%
	End If
%>
        <h2>Scan Glass to Mark Done</h2>
        <fieldset>

         <div class="row">

                        <div class="row">
                <label>Glass ID</label>
                <input type="text" name='ScannedID' id='DoneID' >
            </div>

<script type="text/javascript">

	function callback1(ScannedID) {
		var GlassID = "Glass ID Marked complete:" + ScannedID;

		document.getElementById('ScannedID').innerHTML = GlassID;
		console.log(GlassID);
	}

	function adaptiscanBarcodeFinished(ScannedID, barcodeTypeId, barcodeTypeString) {
		var textbox = document.getElementById("DoneID");
		

		textbox.value = ScannedID;
		igline.submit();
}

</script>

        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>

            </form>

</body>
</html>
