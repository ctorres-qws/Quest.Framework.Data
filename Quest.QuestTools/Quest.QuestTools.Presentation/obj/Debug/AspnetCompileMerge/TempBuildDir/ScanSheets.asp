         <!--#include file="dbpath.asp"-->              
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
           
			<!-- SCAN for Hardware including Gaskets, Angles, Small pieces-->
			<!-- Set up as a basic reader for Barcodes December 2016 -->
			

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Hardware Scan</title>
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



Name = request.Querystring("Name")
Barcode = request.Querystring("Barcode")
Qty = request.Querystring("Qty")
  if Qty = "" then
  Qty = 0
  end if
ScanDATE = DATE

Added = False
Error = "No Error"


'If Statement to check for true scan (Name and Barcode)
If Name <> "" and Barcode <> "" then

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "Select Top 1 * FROM Z_Hardware"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	rs.addnew
	rs.fields("Name") = Name
	rs.fields("Barcode") = Barcode
	rs.fields("Qty") = QTY
	rs.fields("ScanDate") = ScanDate
	Added = TRUE
	
	rs.update
	rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing


else
	if Name = "" and Barcode = "" then
	'Both Empty
		Added = False
		Error = "No Data"
	else 
	' One Field Empty
		Added = False
		Error = "Name or Barcode not included, Please retry"
	end if
end if	



 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Scan" target="_self">Scan Tools</a>
        </div>
   
   
   
     <form id="igline" title="Hardware Scan" class="panel" name="igline" action="ScanHardware.asp" method="GET" selected="true">
      
	  <h2>Hardware Scan</h2>
	  <fieldset>
		 <%
		 if Added = False then
			if Error = "No Data" then
			response.write "<h3>Please Enter a Name and Scan a Barcode:</h3>"
			else
			response.write "<h3>Error: " & Error & "</h3>"
			end if
		 end if
		 if Added = True then
			response.write "<h3>Added " & Name & "</h3>"
			response.write "<h3>Barcode: " & Barcode & "</h3>"
			response.write "<h3>QTY: " & qty & " to Database</h3>"
		 end if
		 %>
		

		<div class="row">
        <label>Name</label>
        <input type="text" name='Name' id='Name' />
        </div>
		
        <div class="row">
        <label>Barcode</label>
        <input type="text" name='Barcode' id='barcode' />
        </div>
    
        <div class="row">
        <label>Qty</label>
        <input type="" name='Qty' id='Qty' value = 1 />
        </div>	
                              	

       <script type="text/javascript">
				  
		function callback1(barcode) {
			var barcodeText = "BARCODE:" + barcode;
			document.getElementById('barcode').innerHTML = barcodeText;
            console.log(barcodeText);
        
            }
            
		function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
		var textbox = document.getElementById("barcode"); 
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