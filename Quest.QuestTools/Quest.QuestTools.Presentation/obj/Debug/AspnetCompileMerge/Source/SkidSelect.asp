<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Select - Allows the selection of a Skid from a list, and then the choice to print -->
<!-- Skid Print - skidprint.asp prints a table view of the skid -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Skid Select</title>
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select distinct name FROM SKIDItem order by name ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection 
%>


    <div class="toolbar">
        <h1 id="pageTitle">Select Skid to View</h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>

 <form id="Skiditem" title="Select QC Inventory Type" class="panel" name="Skiditem" method="GET" target="_self" selected="true">
              
			  <h2>Select Skid</h2>
              <fieldset>               			
			
				<div class="row">
                <label>Skid</label>
					<select name='name='Skids' id='Skids'>
						<%
						rs.movefirst
						do while not rs.eof
							Response.Write "<option value = '"
							Response.Write TRIM(rs("name"))
							Response.Write "'>"
							Response.Write TRIM(rs("name"))
							Response.Write "</option>"
						rs.movenext
						loop
						%>

					</select>
                </div>
				
				
			</fieldset>
        		<a class="whiteButton" onClick="Skiditem.action='SkidPrint.asp'; Skiditem.submit()">Print Skid listing</a><BR>
				<!--<a class="redButton" onClick="QCitem.action='#'; QCitem.submit()">Delete Stock</a><BR>-->
            
         
			
</form>       
  
	<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("Skids");
    
    
        textbox.value = barcode;
    
}
			
			

    

        </script> 
  
  
  
  <%
rs.close
set rs = nothing

DBConnection.close
set DBConnection = nothing
  %>
</body>
</html>
