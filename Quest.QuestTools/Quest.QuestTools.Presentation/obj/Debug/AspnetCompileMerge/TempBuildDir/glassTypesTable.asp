<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Types-->
<!-- Created July 31st, by Michael Bernholtz -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Types / Spacers</title>
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
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job / Colour</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Types / Spacers" selected="true">
		<li><table><tr><td>
<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM XQSU_GlassTypes ORDER BY ID ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
response.write "<table border='1' class='sortable'><tr><th>Type</th><th>Description</th><th>Shop Code</th><th>Status</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Type") & "</td><td>" & RS("Description") &"</td><td>" & RS("ShopCode") &"</td><td>" & RS("Status") &"</td></tr>"
	
	rs.movenext
loop
response.write "</table>"
rs.close
set rs = nothing

%>
</td><td  valign = "top">
<%


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM XQSU_OTSPACER ORDER BY ID ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection
response.write "<table border='1' class='sortable'><tr><th>Spacer</th><th>Overall Thickness</th></tr>"
do while not rs2.eof
	response.write "<tr><td>" & RS2("spacer") & "</td><td>" & RS2("ot") &"</td></tr>"
	
	rs2.movenext
loop
response.write "</table>"

rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing


%>
	</td></tr>

</table></li>	
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
