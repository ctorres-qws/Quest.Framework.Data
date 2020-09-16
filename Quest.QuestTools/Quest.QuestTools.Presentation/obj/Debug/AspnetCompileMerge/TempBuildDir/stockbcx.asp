<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

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
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   



            
            
            
            </form>
            
              <form id="enter" title="Enter Stock" class="panel" name="enter" action="stockbcx.asp#_submit" method="GET" target="_self" selected="true">
              

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

 	po = request.QueryString("Po")
  	die = request.QueryString("die")
    part = request.QueryString("part")
	color = request.QueryString("color")
	length = request.QueryString("length")
    qty = request.QueryString("qty")    

rs.close
set rs = nothing	
%>

        <h2>Scan Stock</h2>
  
                       
                                <fieldset>
  
       <div class="row">
                <label>PO</label>
                <input type="text" name='po' id='Po' value='<% response.write po %>'>
            </div>
 

                        <div class="row">
                <label>Part</label>
                <input type="text" name='partid' id='Partid' value='<% response.write part %>' >
            </div>
            
               <div class="row">
                <label>Supplier Code</label>
                <input type="text" name='die' id='die' value='<% response.write die %>' >
            </div>


                        <div class="row">
                <label>Color</label>
                <input type="text" name='color' id='Color'  value='<% response.write color %>'>
            </div>
            
 
                        <div class="row">
                <label>Length</label>
                <input type="text" name='length' id='Length'  value='<% response.write length %>'>
            </div>
            
              <div class="row">
                <label>Qty</label>
                <input type="text" name='Qty' id='Qty'  value='<% response.write qty %>'>
            </div>
            

                        <div class="row">
                <label>Aisle</label>
                <input type="text" name='aisle' id='Aisle' >
            </div>
            
                           <div class="row">
                <label>Rack</label>
                <input type="text" name='rack' id='Rack' >
            </div>
            
                           <div class="row">
                <label>Shelf</label>
                <input type="text" name='shelf' id='Shelf' >

            </div>
            
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
             <script type="text/javascript">
            
            
         function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    inputs = document.getElementsByTagName('input');
    for (i = 0; i < inputs.length; ++i) {
        if (inputs[i].value == '' || inputs[i].value == 'undefined') {
            inputs[i].value = barcode;
            break;
        }
    }
}

</script>
</fieldset>


        <BR>

            
            </form>

              <form id="submit" title="Enter Stock" class="panel" name="submit" action="stockbc1.asp" method="GET" target="_self">
              
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

 	po = request.QueryString("Po")
  	die = request.QueryString("die")
    part = request.QueryString("part")
	color = request.QueryString("color")
	linch = request.QueryString("length")
    qty = request.QueryString("qty")    
	aisle = request.QueryString("aisle") 
	rack = request.QueryString("rack") 
	shelf = request.QueryString("shelf")   
	
linch = replace(linch,"IN","")
linch = replace(linch," ","")
response.write linch


	
'rs.addnew
'rs.fields("note 2") = po
'rs.fields("note") = die
'rs.fields("part") = part  
'rs.fields("colour") = color 
'rs.fields("linch") = linch 
'rs.fields("qty") = qty 
'rs.fields("aisle") = aisle 
'rs.fields("rack") = rack 
'rs.fields("shelf") = shelf         
'rs.update   

rs.close
set rs = nothing      
%>
                     </form>   
                
 <%
DBConnection.close
Set DBConnection = nothing
%> 
               
</body>
</html>
