<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
<%
Server.ScriptTimeout=300

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' AND COLOUR ='Mill' order by PART ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT PART,DESCRIPTION FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)
	
Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = FixSQL("SELECT JOB FROM Z_Jobs Where Completed = False order by Job ASC")
rs3.Cursortype = GetDBCursorType
rs3.Locktype = GetDBLockType
rs3.Open strSQL3, DBConnection
  
  %>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>
    
      
    
<ul id="screen1" title="Stock Level" selected="true">            

<%
response.write "<li class='group'>Mill Allocation for Stock </li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Mill</th></tr>"

rs2.movefirst
	do while not rs2.eof
	partqty = 0
	rs.filter = "Part = '" & rs2("Part") & "' AND Allocation = 'Stock'" 
	
	do while not rs.eof
					partqty = rs("Qty") + partqty
	rs.movenext
	loop
	rs.filter = ""
	if partqty >0 then
		response.write "<tr><td>" & rs2("part") & "</td><td>" & rs2("description") & "</td><td>" & partqty & "</td></tr>"
	end if
rs2.movenext
loop

response.write "</table></li>"

rs3. movefirst
do while not rs3.eof

response.write "<li class='group'>Mill Allocation for " & rs3("Job") & "</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Mill</th></tr>"

rs2.movefirst
	do while not rs2.eof
	partqty = 0
	rs.filter = "Part = '" & rs2("Part") & "' AND Allocation = '" & rs3("Job") & "'" 
	
	do while not rs.eof
					partqty = rs("Qty") + partqty
	rs.movenext
	loop
	rs.filter = ""
	if partqty >0 then
		response.write "<tr><td>" & rs2("part") & "</td><td>" & rs2("description") & "</td><td>" & partqty & "</td></tr>"
	end if
rs2.movenext
loop

response.write "</table></li>"

rs3.movenext
loop
%>

   
            
   </ul>
</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs=nothing
rs3.close
set rs3=nothing
DBConnection.close
set DBConnection=nothing
%>

