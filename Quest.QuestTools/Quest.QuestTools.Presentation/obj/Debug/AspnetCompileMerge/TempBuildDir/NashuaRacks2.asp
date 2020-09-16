                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Nashua Empty Racks - Report for Shaun and Lev,  April 2017-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Nashua Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    	<% Server.ScriptTimeout = 500 %> 
    <%
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'NASHUA' ORDER BY AISLE ASC, RACK ASC, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Profiles" selected="true">
        
        
<% 
Dim aisle(19)
aisle(1) = "Aa"
aisle(2) = "Bb"
aisle(3) = "Cc"
aisle(4) = "Dd"
aisle(5) = "Ee"
aisle(6) = "Ff"
aisle(7) = "Gg"
aisle(8) = "Hh"
aisle(9) = "Ii"
aisle(10) = "Jj"
aisle(11) = "Kk"
aisle(12) = "Ll"
aisle(13) = "Mm"
aisle(14) = "Nn"
aisle(15) = "Oo"
aisle(16) = "Pp"
aisle(17) = "Qq"
aisle(18) = "Rr"
aisle(19) = "Floor"


shelf= "0"
rack ="0"


CurrentAisle = 1
DO Until CurrentAisle > 19
Rack = 1
Shelf = 1
ShelfFound = "0" 'False
Do Until Rack >8
	Do Until Shelf > 8
		rs.filter =""
		rs.filter = "Aisle = '" & Aisle(CurrentAisle) & "' and Rack = '" & rack & "' AND Shelf =  '" & shelf & "'"
		if rs.eof then
			response.write "<li> Aisle: " & Aisle(CurrentAisle) & " Rack: " & rack & " Shelf: " & shelf & "</li>"
		else
		end if
	Shelf = Shelf + 1
	Loop
Shelf = 1
Rack = Rack +1
Loop 

CurrentAisle = CurrentAisle +1
Loop


rs.close
set Rs = nothing

DBConnection.close
Set DBConnection = nothing
%>
      </ul>                      
               
</body>
</html>
