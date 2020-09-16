<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Item UnFlush - Sets Flush to True and Adds a Flush date -->
<!-- Can be autocalled by scanreport.asp with an add report flag-->


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
 </head>

<body>
     
     <% 
	 
currentDate = Date()

IsError = False
' Reset the Variable for locating an Error
	 
bc = UCASE(request.querystring("barcodeid"))
report = UCASE(request.querystring("report"))
			
			SQL1 = "UPDATE SkidItem Set Flushed = false , FlushedDate = NULL WHERE Flushed = true AND BARCODE = '" & bc & "'"
			Set RS1 = DBCOnnection.Execute(SQL1)
	
DBConnection.close
set DBConnection=nothing
 %>

     


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
   
   <form id="flushskid" title="UnFlush Item" class="panel" name="flushskid" action="skidunflush.asp" target="_self" method="GET" selected="true">
         <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write bc & " - Set to Active<BR>" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Unflush a Skid Item</h2>
        <fieldset>
       
            
                        <div class="row">
                <label>Barcode</label>
                <input type="text" name='barcodeid' id='barcodeid' >
            </div>
            
                              	
            
            
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
        
        
           
         
 
        </fieldset>
        <BR>
            <input type="submit" class="whiteButton" onsubmit=="return confirm('Are you sure you want to do that?');">
	<%		
		 if report = "1" then
			response.write" <a class='whiteButton' target='#_self' href='skidreport.asp'>Back to Report</a>"
		end if
	%>

			</form>



</body>
</html>