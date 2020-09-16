                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Converted to Table Form On August 18th, this is the old row form -->
<!-- Both StockbyRack2 and StockbyRack2Table are options to run-->
<!-- N added to end for Nashua Version Feb 2017-->

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
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    <%
Colour = request.QueryString("Colour")
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'NPREP' AND COLOUR = '" & Colour & "' ORDER BY AISLE, RACK, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection




%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="stockcolorlistPrep.asp" target="_self">Stock</a>
		  </div>
   
   
         
       
        <ul id="Profiles" title="Profiles" selected="true">
        
       <li><li><table border='1' class='sortable'>
		<tr><th>Part</th><th>Qty</th><th>PO</th><th>Bundle</th><th>ExBundle</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>Datein</th></tr>
<% 

do while not rs.eof

Response.write "<TR>"
Response.write "<TD><a href='stockbyrackedit.asp?id=" & rs("id") & "&ticket=PrepColor&colour=" & colour & "' target='_self'>" & rs("part") & "</A></TD>"
Response.write "<TD>" & rs("qty") & "</TD>"
Response.write "<TD>" & rs("po") & "</TD>"
Response.write "<TD>" & rs("bundle") & "</TD>"
Response.write "<TD>" & rs("exbundle") & "</TD>"
Response.write "<TD>" & rs("Aisle") & "</TD>"
Response.write "<TD>" & rs("Rack") & "</TD>"
Response.write "<TD>" & rs("Shelf") & "</TD>"
Response.write "<TD>" & rs("DateIn") & "</TD>"


rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
	</Table></LI>
      </ul>                 
            
       
               
</body>
</html>
