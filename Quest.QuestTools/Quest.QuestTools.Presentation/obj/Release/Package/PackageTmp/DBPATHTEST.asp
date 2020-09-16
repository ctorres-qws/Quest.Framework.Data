                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="dbpath2.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
 
<% PO = Request.querystring("PO") %>

<%
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODE"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From PURCHASES"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection2
%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
   
   
   
    <form id="PrefPurchaseOrder" title="Glass Line Scan" class="panel" action="PREFPOTESTReceive.asp" method="GET" selected="true">

	<% if PO = "" then
		 response.write ""
		 else %>
			<div class="row">
                <% response.write "<label>"& PO & "</Label>" %>
								
			</div>
		<% end if %>	

		
		
		<%
		
		
		do while not rs4.eof
			rs4.movelast
			response.write "<h2> I am from PREF: " & rs4.fields("Number") & "</h2>"
		rs4.movenext
		loop
		do while not rs2.eof
			rs2.movelast
			response.write "<h2> I am from Local DB " & rs2.fields("JOB") & "</h2>"
		rs2.movenext
		loop
		

		%>
		
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Purchase Order</label>
                <input type="text" name='PO' id='iPO' >
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
		PrefPurchaseOrder.submit();
    
}
			
			

    

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:PrefPurchaseOrder.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
