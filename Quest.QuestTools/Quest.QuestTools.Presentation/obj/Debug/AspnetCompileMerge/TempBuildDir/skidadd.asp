<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Add - allows an item to scanned by barcode or entered manually -->
<!-- Scanning an item that starts with SKID automatically goes to skid field -->
<!-- Scanning an item that starts with FLUSH automatically jumps to the flush page and flushes the skid (using an Add flag)-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Scan to Skid</title>
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

currentDate = Date()


bc = TRIM(UCASE(request.querystring("barcodeid")))
skidname = TRIM(UCASE(request.querystring("skidname")))

IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found on the Z_GLASSDB


if skidname = "" or bc = "" then
else
	'Check if Glass Backorder (Only Glass)
	'For now assume glass if Backorder - Use the Z_GLASSDB to get JOB FLOOR TAG and add item to ScanItem
	'If not backorder jump down to plain entry from Valid Barcode

	' GT code means Backorder, so the code will run first, if the barcode is not a backorder, go down to the else
	' There is a patch in place, GT is code for BackOrder, so GTM had to be hardcoded or it ran like a backorder
	if Left(bc,2) = "GT" AND NOT Left(bc,3) ="GTM" then
		if len(bc) <3 then 
			bc = "GT00"
		end if	

	'Drop the GT in front of the ID to get the ID number
		GlassID = Mid(bc, 3)

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

		Do while not rs.eof
			if CLng(rs("ID")) = CLng(GlassID) then
			
					' Details of Completed Item to be shown
					BARCODE = bc
					JOB = rs.fields("NOTE 1")
					FLOOR = rs.fields("NOTE 2")
					TAG = rs.fields("NOTE 3")
					Descriptor = "GLASS"
		
		'After collecting the Backorder Data insert it into the SkidItem if it is not already there
		
				Set rs2 = Server.CreateObject("adodb.recordset")
				strSQL2 = "SELECT * FROM SKIDITEM WHERE FLUSHED = False ORDER BY ID ASC"
				rs2.Cursortype = 2
				rs2.Locktype = 3
				rs2.Open strSQL2, DBConnection
				rs2.filter = "BARCODE = '" & bc & "'"
				if rs2.bof then  
				' Check to see if the filter finds an existing barcode - if not ADD else - already added
						rs2.addnew 
						rs2.fields("BARCODE") = bc
						rs2.fields("name") =  skidname
						rs2.fields("JOB") = JOB
						rs2.fields("FLOOR") = FLOOR
						rs2.fields("TAG") = TAG
						rs2.fields("DESCRIPTOR") = Descriptor
						rs2.fields("SCANDATE") = currentDate
						rs2.UPDATE		
				else 
				' Filter on Barcode already found a match, so the item has already been scanned
					error = " Barcode: " & bc & " already Scanned."
					IsError = True
				end if	
				IDFound = True	
			end if
		rs.movenext
		loop
		
		if IDFound = False then 
		' Backorder Record not found in Z_GLASSDB for comparison
			error = " Scanned ID: " & GlassID & " does not exist as a Backorder."
			IsError = True
		end if
		
	else
	' If the Glass is not a Backorder - Read the Barcode to find out JOB FLOOR TAG DESCRIPTOR

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
		
		
		Descriptor = " N/A" 
		' Collect the Barcode Data from the Scanned Barcode

			Set rs2 = Server.CreateObject("adodb.recordset")
			strSQL2 = "SELECT * FROM SKIDITEM where BARCODE ='" & bc & "' AND FLUSHED = False ORDER BY ID ASC"
			rs2.Cursortype = 2
			rs2.Locktype = 3
			rs2.Open strSQL2, DBConnection

			if rs2.bof then  
				if Len(bc) > 3 then

					' Check that Barcode is longer than three digits in length
					rs2.addnew 
					rs2.fields("BARCODE") = bc
					rs2.fields("JOB") = jobname
					rs2.fields("FLOOR") = floor
					rs2.fields("TAG") = tag
					rs2.fields("DESCRIPTOR") = Descriptor
					rs2.fields("SCANDATE") = currentDate
					rs2.fields("name") = skidname
					rs2.UPDATE		
				else 
					error = bc & ": Not a Valid Barcode, Try Again"
					IsError = True
				end if		
					
					
			else 

			' Filter on Barcode already found a match, so the item has already been scanned
				error = " Barcode: " & bc & " already Scanned."
				IsError = True
			end if	
			IDFound = True	
		end if
	end if
  

 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
   
   <form id="AddSkidItem" title="Add to Skid Scan" class="panel" name="AddSkidItem" action="skidadd.asp" method="GET" selected="true">
         <% if bc = "" or skidname = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = false then
				response.write bc & " added to Skid. <BR>" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
		end if %>
        <h2>Scan Item to Skid (or Skid Flush Marker)</h2>
        <fieldset>
       
            
         <div class="row">
                <label>Barcode</label>
                <input type="text" name='barcodeid' id='barcodeid' >
        </div>
 


         <div class="row">
                <label>Skid</label>
                <input type="text" name='skidname' id='skidname' value ="<%response.write trim(skidname)%>">
        </div>
                              	
            
            
    <script type="text/javascript">
				  
  
	  function callback1(barcode) {
        var barcodeText = "BARCODE:" + barcode;

        document.getElementById('barcode').innerHTML = barcodeText;
        console.log(barcodeText);
        
    }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("barcodeid");

	if ( barcode.substring(0,4) == "SKID")  {
		barcode = barcode.substring(4, barcode.length);
		textbox = document.getElementById("skidname");
	
    }
	if ( barcode.substring(0,5) == "FLUSH")  {
		barcode = barcode.substring(5, barcode.length);
		window.location.replace("SkidFlush.asp?add=1&skidname=" + barcode);
    }
	
	
    textbox.value = barcode;
    
}
        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:AddSkidItem.submit()">Submit</a>
            
            
            
            </form>

			
<%
' Need to add the correct recordsets to close
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
set DBConnection=nothing
%>	

</body>
</html>