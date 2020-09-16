<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Search Barcode  Displays the skid where a specific item is located-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Skid</title>
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
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
   
    <form id="SearchSkid" title="Search Skids" class="panel" name="SearchSkid" action="SkidSearchResults.asp" method="GET" selected="true">
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
        <h2>Skid Scan</h2>
        <fieldset>
       
            
         <div class="row">
                <label>Barcode</label>
                <input type="text" name='barcodeid' id='barcodeid' >
        </div>
 


         <div class="row">
                <label>Job</label>
                <input type="text" name='job' id='job' >
        </div>
		
		 <div class="row">
                <label>Floor</label>
                <input type="text" name='floor' id='floor' >
        </div>
				
		 <div class="row">
                <label>Tag</label>
                <input type="text" name='tag' id='tag' >
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
        <a class="whiteButton" href="javascript:SearchSkid.submit()">Submit</a>
            
            
            
            </form>
			
<%
' Need to add the correct recordsets to close
rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>	

</body>
</html>