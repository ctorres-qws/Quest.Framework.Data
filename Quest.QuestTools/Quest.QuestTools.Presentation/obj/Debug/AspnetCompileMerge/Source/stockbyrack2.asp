<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Converted to Table Form On August 18th, this is the old row form -->
<!-- Both StockbyRack2 and StockbyRack2Table are options to run-->

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
	
Set rs = Server.CreateObject("adodb.recordset")
If CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE='JUPITER' ORDER BY AISLE ASC, RACK ASC, SHELF DESC"
else
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE='GOREWAY' ORDER BY AISLE ASC, RACK ASC, SHELF DESC"
end if
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

afilter = request.QueryString("aisle")


%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Stock by Rack</h1>
        <a class="button leftButton" type="cancel" href="stockbyaisle.asp" target="_self">Stock</a>
		  </div>
   
   
         
       
        <ul id="Profiles" title="Profiles" selected="true">
        <li class='group'><a href="stockbyrack2table.asp?aisle=<% response.write afilter %>" target="_self"> Switch to Table Form - Aisle <% response.write afilter %></a></li>
      
        
<% 

do while not rs.eof
if rs("aisle") = afilter then
part = rs("part")
qty = rs("qty")
id = rs("ID")
po = rs("PO")
bundle = rs("Bundle")
shelf = rs("shelf")
colour = rs("colour")
datein = rs("datein")
	if aisle = rs("aisle") then
	else
	response.write "<li class='group'>Aisle " & rs("aisle") & "</li>"
	end if
	
	if rack = rs("rack") then
	else
	if ISNULL(rack) = -1 then
	else
	response.write "<li class='group'>Rack " & rs("rack") & "</li>"
	end if
	end if
	

%>
<!-- At request of Ruslan, Added Colour to this screen, January 17 2014, Michael Bernholtz-->
<li><a href="stockbyrackedit.asp?id=<% response.write id %>&aisle=<% response.write afilter %>&ticket=GOREWAY" target="_self">Shelf <%response.write shelf & ", " & part & ", " & qty & " SL" & " PO " & po & " /" & bundle & " - " & colour & " Entered: " & datein %></a></li>
<%
end if

aisle = rs("aisle")
rack = rs("Rack")
rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
