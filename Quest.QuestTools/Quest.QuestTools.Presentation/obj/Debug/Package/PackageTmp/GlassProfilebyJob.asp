<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Job Search Glass Profiles</title>
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
	
	JOB = Request.querystring("JOB")
	
	
Set rs = Server.CreateObject("adodb.recordset")
if UCASE(JOB) = "ALL" then
	strSQL = "SELECT * FROM GlassTypes ORDER BY JOB, NAME ASC"
else
	strSQL = "SELECT * FROM GlassTypes WHERE [JOB] = '" & JOB & "' ORDER BY NAME ASC"
end if

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
       <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Panel</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Profiles - Job Search" selected="true">
<% 
response.write "<li class='group'>Job Search of Styles - " & JOB & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Name</th><th>Description</th><th>Job</th><th>Side</th><th>Material</th><th>Colour</th><th>Offset X</th><th>OffSet Y</th><th>Notes</th><th>Top Left</th><th>Top Right</th><th>Bottom Right</th><th>Bottom Left</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Name") & "</td><td>" & RS("Description") & "</td><td>" & RS("JOB") &"</td><td>" & RS("Side") & "</td><td>" & RS("MAterial") & "''</td><td>" & RS("Colour") & "''</td><td>" & RS("OffsetX") & "</td><td>" & RS("OffsetY") & "</td><td>" & RS("Notes") & "</td>" 
	
	AValue = RS("A1") & RS("A2") & RS("A3") & RS("A4") & RS("A5") & RS("A6") & RS("A7") & RS("A8")
	BValue = RS("B1") & RS("B2") & RS("B3") & RS("B4") & RS("B5") & RS("B6") & RS("B7") & RS("B8")
	CValue = RS("C1") & RS("C2") & RS("C3") & RS("C4") & RS("C5") & RS("C6") & RS("C7") & RS("C8")
	DValue = RS("D1") & RS("D2") & RS("D3") & RS("D4") & RS("D5") & RS("D6") & RS("D7") & RS("D8")
	response.write "<td>" & AValue & "</td><td>" & BValue & "</td><td>" & CValue & "</td><td>" & DValue & "</td>"
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
