<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Sheet Inventory Page - Later will be broken down like Inventory report-->
<!-- Created March 2017, by Michael Bernholtz, Y_SHEET_INV-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Sheet Inventory</title>
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
strSQL = "SELECT * FROM Y_SHEET_INV ORDER BY ID ASC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Panels</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Sheet Report" selected="true">
        
        
<% 
response.write "<li class='group'>Sheet REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Material</th><th>Thickness</th><th>Entry Date</th><th>Qty</th><th>Last Modified</th><th>PO</th><th>Location</th><th>Entry Qty</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Material") & "</td>"
	response.write "<td>" & RS("Thickness") & "</td>"
	response.write "<td>" & RS("EntryDate") & "</td>"
	response.write "<td>" & RS("Qty") & "</td>"
	response.write "<td>" & RS("LastModify") & "</td>"
	response.write "<td>" & RS("PO") & "</td>"
	response.write "<td>" & RS("Location") & "</td>"
	response.write "<td>" & RS("EntryQty") & "</td>"
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li>"



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
               
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
