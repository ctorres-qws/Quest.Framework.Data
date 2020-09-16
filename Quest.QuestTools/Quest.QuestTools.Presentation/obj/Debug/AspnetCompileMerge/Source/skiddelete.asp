<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Remove - Flushes a specific item in the table -->
<!-- Skid Delete - Deletes a specific item from the record -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Scan Remove from Skid</title>
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


bc = UCASE(request.querystring("barcodeid"))

IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found in SkidItem
Result = ""
' Namespace to hold deleted information to display on correct Removal

if bc = "" then
else
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM SKIDITEM ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	rs.filter = "BARCODE = '" & bc & "'"
	
	if not rs.bof then
			Result = bc & " Removed from Skid " & rs("name") 
			SQL1 = "Delete from SkidItem WHERE Barcode = '" & bc & "'"
			Set RS1 = DBCOnnection.Execute(SQL1)

		IDFound = True
	else
	
		IsError = True
		error = bc & " not found on a Skid"
	end if
end if	
  

 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
   
   <form id="DeleteSkidItem" title="Delete from Skid Scan" class="panel" name="DeleteSkidItem" action="skiddelete.asp" method="GET" selected="true">
         <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = false then
				response.write Result & "<BR>" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
		end if %>
        <h2>Skid Scan</h2>
        <fieldset>
       
            
         <div class="row">
                <label>Barcode</label>
                <input type="text" name='barcodeid' id='barcodeid' >
        </div>
 

 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:DeleteSkidItem.submit()">Submit</a>
            
            
            
            </form>
    <script type="text/javascript">
				  
				 		  
    function callback1(barcode) {
        var barcodeText = "BARCODE:" + barcode;

        document.getElementById('barcode').innerHTML = barcodeText;
        console.log(barcodeText);
        
    }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("barcodeid");
  
    
    textbox.value = barcode;
    
}
	
        </script>
			
<%
' Need to add the correct recordsets to close
rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>	

</body>
</html>