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
   
  <%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_Master order by part ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_Color Order by Project ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection  
   %>
   
    <form id="part1" title="Stock by Die" class="panel" name="part1" action="stockbydie.asp" method="GET" target="_self" >
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%



var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART")
response.write ""
end if
part = rs("PART")
rs.movenext

loop
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:part1.submit()">Submit</a>
            
            
            
            </form>
            
              <form id="enter" title="Enter Stock" class="panel" name="enter" action="stockin.asp" method="GET" target="_self" selected="true">
              
                              


        <h2>Select Die</h2>
  
                       
                                <fieldset>
  <div class="row">
     <label>Select Die</label>
<select name="part">

<option selected name=part value="-">-<%



var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART")
response.write ""
end if
part = rs("PART")
rs.movenext

loop
%></select>
</div>

   <div class="row">
<!-- Colour Edited to be a Drop-Down from the Y_Color table - At Request of Ruslan - Michael Bernholtz, January 20, 2014-->
            <div class="row">
             <label>Color</label>
            <select name="color" id='color' >
<%

rs2.movefirst
Do While Not rs2.eof

Response.Write "<option name=color value='"
Response.Write rs2("Project")
Response.Write "'>"
Response.Write rs2("Project")
response.write "</option>"

rs2.movenext

loop
%></select></DIV>

                        <div class="row">
                <label>Length</label>
                <input type="text" name='length' id='length' >
            </div>
            
              <div class="row">
                <label>Qty</label>
                <input type="text" name='Qty' id='Qty' >
            </div>
            
                        <div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO' >
            </div>
            
            
                     <div class="row">

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
            
            <div class="row">
             <label>Warehouse</label>
            <select name="warehouse">
<option selected name=jobname value="-">-<%

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM Y_WAREHOUSE"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

varw = 0
rs3.movefirst
Response.Write "<option SELECTED=jobname value='"
Response.Write rs3("NAME")
Response.Write "'>"
Response.Write rs3("NAME")
response.write ""
RS3.MOVENEXT

Do While Not rs3.eof

Response.Write "<option name=jobname value='"
Response.Write rs3("NAME")
Response.Write "'>"
Response.Write rs3("NAME")
response.write ""

rs3.movenext

loop
%></select></DIV>
            
                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
         
</fieldset>

<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("color");
    
    
        textbox.value = barcode;
	    
}
</script>

        <BR>

            
            </form>
            
             <form id="remove" title="Remove / Deplete / Location" class="panel" name="remove" action="stockedit.asp" method="GET" target="_self">
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART")
response.write ""
end if
part = rs("PART")
rs.movenext

loop
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:remove.submit()">Submit</a>
            
            
            
            </form>
                
<select name="COLOUR">
<option selected name=jobname value="-">-<%
var1 = 0
rs2.movefirst
Do While Not rs2.eof

If rs2("COLORCODE") = COLORCODE then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs2("COLORCODE")
Response.Write "'>"
Response.Write rs2("COLORCODE")
response.write ""
end if
COLORCODE = rs2("COLORCODE")


rs2.movenext
loop
%>
</select>                

<%

rs.close
set rs = nothing
rs2.close
set rs2= nothing
rs3.close
set rs3 = nothing
DBConnection.close
set DBConnection= nothing
%>                
             
               
</body>
</html>
