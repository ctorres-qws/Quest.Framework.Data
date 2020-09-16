<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodega.asp"-->
			<!-- Page Created December 12th by Michael Bernholtz on request of Jody Cash-->
			<!-- Adaptiscan GT:ID to unmark an item Complete and reset to Active  -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Re-Activate</title>
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
			if isDate(rs.fields("COMPLETEDDATE")) = True then
				' Record in Z_GLASSDB matches the scanned item and has an Output Date
				' Successfully Marked Active
				rs.fields("COMPLETEDDATE") = NULL
				' Details of Completed Item to be shown
				DetailsID = rs.fields("JOB") & " " & rs.fields("FLOOR") & "  " & rs.fields("TAG")
				
				' Delete the Record from the X_BARCODEGA
				Set rs2 = Server.CreateObject("adodb.recordset")
				strSQL2 = "DELETE * FROM X_BARCODEGA WHERE BARCODE = '" & ScannedID & "' "
				rs2.Cursortype = 2
				rs2.Locktype = 3
				rs2.Open strSQL2, DBConnection
			else 	
				' Record in Z_GLASSDB matches the scanned item  but already has no Output Date
				' Previously Marked Active
				error = "Glass Order: " & GlassID & " already marked active." 
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
   
   
   
    <form id="igline" title="Glass Line Scan" class="panel" name="igline" action="glassReactivateScan.asp" method="GET" selected="true">
         <% if ScannedID = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write GlassID & " - Marked Active: " & DetailsID
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Scan Glass to Mark Active</h2>
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
