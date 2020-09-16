<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
			<!-- UNSCAN code for DEPTVAR = Forel   - Resets the Optima Date, Completed Date, and Department -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass UnScan</title>
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
error = ""

bc = UCASE(request.querystring("window"))

%>
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

DEPTVAR = "Forel"

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

If Len(bc) > 0 Then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

End If

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEGA WHERE Barcode='" & bc & "'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_EMPLOYEES"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

' GT code means Backorder, so the code will run first, if the barcode is not a backorder, go down to the else
' There is a patch in place, GT is code for BackOrder, so GTM had to be hardcoded or it ran like a backorder
if Left(bc,2) = "GT" AND NOT Left(bc,3) ="GTM" then
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

			Do while not rs.eof
				if rs.fields("Barcode") = bc and rs.fields("DEPT")=DEPTVAR then
					RecordLocated = rs.Fields("ID")
					rs.fields("DEPT") = "UNSCAN"
					rs.update
					rs4.fields("CompletedDate") = "UNSCAN"
					rs4.fields("OptimaDate") = "UNSCAN"
					rs4.update
				end if
			rs.movenext
			loop
								
				'CHeck to see if the record already exists in X_BARCODEGA and has DEPTVAR as Department
				' If it does, then just change the Department to UNSCAN and mark the current Z_GLASSDB AS Not complete 
				' If it does not, then report that there is no record to change
				' RecordLocated = ID
				
			If RecordLocated = 0 then

			
				' Record in Z_GLASSDB matches the scanned item  but already has a different DEPTVAR

				error = "Glass Order: " & GlassID & " already marked UNSCAN." 
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

If Not rs.eof Then
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



	glasstype = right(bc,2)


	sizecheckid = 0
	
	rs.movefirst
	Do while not rs.eof
		if rs("BARCODE") = bc AND rs("DEPT") = "UNSCAN" then
			error = bc & ": Already UnScanned"
			IsError = True
		end if	
		
		
		if rs("BARCODE") = bc AND rs("DEPT") = DEPTVAR then

			rs.fields("DEPT") = "UNSCAN"
			rs.UPDATE
		end if
	
	rs.movenext
	loop
End If

end if

DbCloseAll

End Function

 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Forel UnScan" class="panel" name="igline" action="ScanGlassForelUnscan.asp" method="GET" selected="true">
         <h2>Forel Glass Line UnScan</h2>
		 <h2>Removes Broken Glass! </h2>
		 <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write bc & " - UnScanned" 
				' & DetailsID
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
            
                              	

            
            
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
			
<%
'rs.close
'set rs=nothing
'rs2.close
'set rs2=nothing
'rs4.close
'set rs4=nothing
'DBConnection.close
'set DBConnection=nothing
%>			
			
</body>
</html>