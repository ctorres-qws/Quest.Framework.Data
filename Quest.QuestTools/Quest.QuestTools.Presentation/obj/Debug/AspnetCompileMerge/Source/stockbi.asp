                       
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
   
   
   
    <form id="part" title="Stock by Die" class="panel" name="part" action="stockbydie.asp" method="GET" target="_self" selected="true">
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_INVLOG"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

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

rs.close
set rs = nothing
rs2.close
set rs2 = nothing 
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:part.submit()">Submit</a>
            
            
            
            </form>
            
            
            
              <form id="enter" title="Enter Stock" class="panel" name="enter" action="stockin.asp" method="GET" target="_self" >
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

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
rs.close
set rs = nothing
%></select>

</div></fieldset><fieldset>


  <div class="row">

                        <div class="row">
                <label>Color</label>
                <input type="text" name='color' id='color' >
            </div>
            
         <div class="row">

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

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

varw = 0
rs2.movefirst
Response.Write "<option SELECTED=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""
RS2.MOVENEXT

Do While Not rs2.eof

Response.Write "<option name=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""

rs2.movenext

loop

rs2.close
set rs2 = nothing
%></select></DIV>

</fieldset>



        <BR>
        <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
            
            
            
            </form>
            
             <form id="remove" title="Edit Details" class="panel" name="remove" action="stockedit.asp" method="GET" target="_self">
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_COLOR"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

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

rs.close
set rs = nothing
rs2.close
set rs2= nothing
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:remove.submit()">Submit</a>
            
            
            
            </form>
            
            <form id="del" title="Delete Stock!" class="panel" name="del" action="stockdel.asp" method="GET" target="_self">
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_COLOR"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

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
rs.close
set rs = nothing
rs2.close
set rs2= nothing
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:del.submit()">Submit</a>
            
            
            
            </form>
			
			
			<%
			DBConnection.close
			Set DBConnection = nothing
             %>  
               
                
             
               
</body>
</html>
