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
 <%
inventorytype = Request.Querystring("inventorytype")
if inventorytype="" then
	inventorytype = "Extrusion"
End if
%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
   
    <form id="part" title="Stock by Die" class="panel" name="part" action="stockbydie.asp" method="GET" target="_self" >
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="part">
<option selected name=jobname value="-">-<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
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
Response.Write rs("PART") & " (" & rs("Description") & ")"
response.write "</option>"
end if
part = rs("PART")
rs.movenext

loop

rs.close
set rs=nothing
rs2.close
set rs2 = nothing
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:part.submit()">Submit</a>
            
            
            
            </form>
            
            
            
              <form id="enter" title="Enter Stock" class="panel" name="enter" action="stockin.asp" method="GET" target="_self" selected="true">
        <h2>Select Die</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->


<select name="part">
<option selected name=jobname value="-">-</option>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER  ORDER BY PART"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.movefirst
inventorytype = REQUEST.QUERYSTRING("inventorytype")
if inventorytype = "" then
else
rs.filter = "inventorytype = '" & inventorytype & "'"
end if

var1 = 0

Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART") & " (" & rs("Description") & ")"
response.write "</option>"
end if
part = rs("PART")


rs.movenext

loop
rs.close
set rs=nothing

%></select>

</div></fieldset><fieldset>


  <div class="row">
<!-- Colour Edited to be a Drop-Down from the Y_Color table - At Request of Ruslan - Michael Bernholtz, January 20, 2014-->
            <div class="row">
             <label>Color</label>
            <select name="color" id='color' >
<%			
' Special Type for Hardware			
			if inventorytype = "Hardware" then		
				response.write "<option value='Hardware' selected >Hardware</option> "
			end if


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_Color WHERE ACTIVE = TRUE Order by PROJECT ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

rs2.movefirst
Do While Not rs2.eof

Response.Write "<option name='color' value='"
Response.Write rs2("Project")' & " - " & rs2("DESC") & " - " & rs2("SIDE")
Response.Write "'"
IF rs2("Project") = "Mill" then
Response.Write " selected"
End IF
Response.Write ">"
Response.Write rs2("Project") & " - (" & rs2("Code") & ")" '& " - " & rs2("SIDE")
response.write "</option>"

rs2.movenext

loop

			if inventorytype = "Sheet" then		
				response.write "<option value='Mill' selected >Mill</option> "
				response.write "<option value='Net Mill'>Net Mill</option> "
				response.write "<option value='Stainless Steel' >Stainless Steel</option> "
			end if

rs2.close
set rs2 = nothing
%></select></DIV>
				<!--   Allocation is not specific enough at Mill Entry, this has been removed - Shaun September 2014, reAdded December 2014-->
			<div class="row">
				<label>Allocation</label>
				<select name="Allocation">
					<% ActiveOnly = True %>
					<option value="" > -  </option>
					<option value="White" >White</option>
					<!--#include file="JobsList.inc"-->
				</select>
				<%
				rsJob.close
				set rsJob=nothing
				%>
			</div>
			<%
			if inventorytype = "Sheet" then		
			%>
            <div class="row">
                <label>Thickness</label>
				<select name="Thickness">
					<% ActiveOnly = True %>
					<option value= "0.27" >0.27 </option>
					<option value= "0.50" >0.50 </option>
					<option value= "0.63" >0.63 </option>
					<option value= "0.80" >0.80</option>
					<option value= "0.125" >0.125</option>
				</select>
            </div>
			<%
			else
			%>
			  <div class="row">
                <label>Length</label>
                <input type="text" name='length' id='length' >
            </div>
			<%
			end if 
			%>
			
            
            <div class="row">
                <label>Qty</label>
                <input type="text" name='Qty' id='Qty' >
            </div>
            
            <div class="row">
                <label>Bundle</label>
                <input type="text" name='Bundle' id='Bundle' >
            </div>
			
			<div class="row">
                <label>Ext. Bundle</label>
                <input type="text" name='ExBundle' id='ExBundle' >
            </div>
            
			<div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO' >
            </div>
			<div class="row">
                <label>Colour PO</label>
                <input type="text" name='ColorPO' id='ColorPO' >
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
			
			<div class="row"> <!-- Date Field Added April 2014 - also updated in Stockin-->
                <label>Expected Date</label>
                <input type="date" name='expdate' id='expdate' >	
				
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


rs2.movefirst
Response.Write "<option SELECTED=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
Response.write "</option>"
RS2.MOVENEXT

Do While Not rs2.eof

Response.Write "<option name=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
Response.write "</option>"

rs2.movenext

loop

rs2.close
set rs2 = nothing
%></select></DIV>
<input type="hidden" name="inventorytype" id="inventorytype" Value ="<%response.write inventorytype%>" />
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
strSQL = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Commented out - Not in use - Older version may have used, do not want to delete
'Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL = "SELECT * From Y_COLOR"
'rs2.Cursortype = 2
'rs2.Locktype = 3
'rs2.Open strSQL, DBConnection

var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART") & " (" & rs("Description") & ")"
Response.write "</option>"
end if
part = rs("PART")
rs.movenext

loop
rs.close
set rs=nothing
'rs2.close
'set rs2 = nothing
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
strSQL = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Commented out - Not in use - Older version may have used, do not want to delete
'Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL = "SELECT * From Y_COLOR"
'rs2.Cursortype = 2
'rs2.Locktype = 3
'rs2.Open strSQL, DBConnection

var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("part") = part then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("PART")
Response.Write "'>"
Response.Write rs("PART") & " (" & rs("Description") & ")"
Response.write "</option>"
end if
part = rs("PART")
rs.movenext

loop
rs.close
set rs=nothing
'rs2.close
'set rs2 = nothing
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:del.submit()">Submit</a>
            
            
            
            </form>
                
             <!--Added January 17, 2014 - Michael Bernholtz for Ruslan - Stock by Color-->
			 <!--To implement this, need to match up to Master Colour-->
				<form id="colour" title="Edit Details" class="panel" name="Colour" action="stockeditcolour.asp" method="GET" target="_self">
        <h2>Select Master Colour</h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="colour">
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_Color ORDER BY Project ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

var1 = 0
rs.movefirst
Do While Not rs.eof

If rs("project") = colour then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs("project")
Response.Write "'>"
Response.Write rs("project")
Response.write "</option>"
end if
colour = rs("project")
rs.movenext

loop
rs.close
set rs=nothing
%></select>
            </div></fieldset>
        <BR>
        <a class="whiteButton" href="javascript:colour.submit()">Submit</a>
            
            
            
            </form>			 
                
 <%
DBConnection.close
set DBConnection = nothing
%>            
               
</body>

</html>
