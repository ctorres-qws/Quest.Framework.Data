                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodega.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Scan</title>
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
	 
'employeeID = 0
'employeeID = request.QueryString("EmployeeID")
'window = request.QueryString("Window") 
  

STAMP = REQUEST.QueryString("STAMP")

STAMPVAR = year(now) & " " & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

'bc = request.querystring("barcode")

DEPTVAR = "GLASSLINE"




bc = request.querystring("window")

OUTDATE = DATE
'The date to add on the item marked completed
IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found on the Z_GLASSDB
DetailsID =""
' Details to show for item successfully marked complete


if Left(bc,2) = "GT" then
	if len(bc) <3 then 
		bc = "GT00"
	end if	

 'Drop the GT in front of the ID to get the ID number
	GlassID = Mid(bc, 3)

	Set rs4 = Server.CreateObject("adodb.recordset")
	strSQL4 = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
	rs4.Cursortype = 2
	rs4.Locktype = 3
	rs4.Open strSQL4, DBConnection


	Do while not rs4.eof
		if CLng(rs4("ID")) = CLng(GlassID) then
			if isDate(rs4.fields("Note 18")) = False then
				' Record in Z_GLASSDB matches the scanned item and does not have an Output Date
				' Successfully Marked Done
				rs4.fields("Note 18") = OUTDATE
				' Details of Completed Item to be shown
				JOB = rs4.fields("NOTE 1")
				FLOOR = rs4.fields("NOTE 2")
				TAG = rs4.fields("NOTE 3")
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
				'Add a Record to X_BARCODEGA code
				'rs = "Select * FROM X_BARCODEGA" from include
				
				rs.addnew 
				rs.fields("BARCODE") = bc
				rs.fields("JOB") = JOB
				rs.fields("FLOOR") = FLOOR
				rs.fields("TAG") = TAG
				rs.fields("DEPT") = DEPTVAR
				rs.fields("DATETIME") = STAMPVAR
				rs.fields("TYPE") = GLASSTYPE
				rs.fields("DAY") = cday
				rs.fields("MONTH") = cmonth
				rs.fields("YEAR") = cyear
				rs.fields("WEEK") = weeknumber
				rs.fields("TIME") = cctime
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
  
	if IDFound = False then 
		' Record not found in Z_GLASSDB for comparison
		error =" Scanned ID: " & GlassID & " does not exist."
		IsError = True
	end if

else


	jobname = Left(bc, 3)
	if inStr(1, bc, "-", 0) = 5 then
		floor = Mid(bc, 4, 1)
		tag = Mid(bc, 5, 5)
	END IF

	if inStr(1, bc, "-", 0) = 6 then
		floor = Mid(bc, 4, 2)
		tag = Mid(bc, 6, 5)
	end if

	if inStr(1, bc, "-", 0) = 7 then
		floor = Mid(bc, 4, 3)
		tag = Mid(bc, 7, 5)
	end if



	glasstype = right(bc,2)


	sizecheckid = 0

	Do while not rs.eof
		if rs("BARCODE") = bc AND rs("DEPT") = "GLASSLINE" then
			sizecheckid = rs("ID")
			error = bc & ": Already Scanned - Not Sent"
			IsError = True
		end if
	rs.movenext
	loop
  
	if sizecheckid = 0 then

		if Len(bc) > 3 then

			rs.addnew 
			rs.fields("BARCODE") = bc
			rs.fields("JOB") = jobname
			rs.fields("FLOOR") = floor
			rs.fields("TAG") = tag
			rs.fields("DEPT") = DEPTVAR
			rs.fields("DATETIME") = STAMPVAR
			rs.fields("TYPE") = glasstype
			rs.fields("DAY") = cday
			rs.fields("MONTH") = cmonth
			rs.fields("YEAR") = cyear
			rs.fields("WEEK") = weeknumber
			rs.fields("TIME") = cctime
			rs.UPDATE
			DetailsID = "JOB: " &jobname & " FLOOR: " & floor & " TAG: " & tag

		else 
			error = bc & ": Not a Valid Barcode, Try Again"
			IsError = True
		end if

	end if
end if
  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Glass Line Scan" class="panel" name="igline" action="ga1.asp" method="GET" selected="true">
         <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write window & " - Sent" 
				' & DetailsID
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Scan Glass</h2>
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
            
                              	
                <% 'response.write window & "<Br>" & employeeID %>
            
            
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
</body>
</html>
