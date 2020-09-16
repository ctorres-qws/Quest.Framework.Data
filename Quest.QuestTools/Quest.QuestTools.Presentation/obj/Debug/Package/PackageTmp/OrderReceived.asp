                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- List of all completed Orders - to follow Order List - on GlassOrder Table
<!-- Designed for Michael Angel and Lev Bedeov-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Orders Received</title>
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
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM GlassOrder WHERE Received = TRUE ORDER BY PO ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection




%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Order" target="_self">Order Entry</a>
        </div>
   
   
         
       
        <ul id="Profiles" title=" Glass Report - All Active" selected="true">
    
<% 

response.write "<li>Orders Marked Received</li>  "
response.write "<li><table border='1' class='sortable' ><thead><tr><th>PO</th><th>Glass Code</th><th>Job</th><th>Floor</th><th>Quantity</th><th>From</th><th>Order By</th><th>ShipQT</th><th>Order</th><th>Expected</th><th>Received</th><th>Notes</th><th>Acknowledged</th><th>Broken</th><th># Broken</th></th></tr></thead><tbody>"
if rs.eof then
Response.write "<tr><td colspan ='14'>No current orders</td></tr>"
end if	
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("PO") & "</td>"
	response.write "<td>" & RS("GlassCode") &"</td>"
	response.write "<td>" & RS("JOB") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("Qty") & "</td>"
	response.write "<td>" & RS("From") & "</td>"
	response.write "<td>" & RS("OrderBy") & "</td>"
	response.write "<td>" & RS("ShipOutDate") & "</td>"
	response.write "<td>" & RS("OrderDate") & "</td>"
	response.write "<td>" & RS("ExpectedDate") & "</td>"
	response.write "<td>" & RS("ReceivedDate") & "</td>"
	response.write "<td>" & RS("Notes") & "</td>"
	response.write "<td>" & RS("Ack") & "</td>"
	response.write "<td>" & RS("Broken") & "</td>"
	response.write "<td>" & RS("Return") & "</td>"
	response.write " </tr>"

	rs.movenext
loop
response.write "</tbody></table>"


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing



%>
      </ul>                 
            
            
            
       
            
              
               
                
             
               
</body>
</html>
