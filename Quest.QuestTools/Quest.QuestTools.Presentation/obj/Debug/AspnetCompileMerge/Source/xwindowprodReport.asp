                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--X Window Production collects many different facets of information Floor wide for each project - This report shows it (until now it remained mostly hidden)-->
<!-- X_Win_PROD table displayed in table form-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>X Window Prod Summary</title>
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
strSQL = "SELECT * FROM X_WIN_PROD ORDER BY JOB, FLOOR ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">X Window Prod Summary</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Reports" target="_self">Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="X-Wing Summary Page" selected="true">
<% 
response.write "<li class='group'>X Window Production Summary</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>Job</th><th>Floor</th><th>Date</th><th>Cycles</th><th>Total Sqft</th><th>Windows</th>"
response.write "<th>AWG</th><th>Panels</th><th>SP</th><th>SU Small</th><th>SU Medium</th><th>SU Large</th>"
response.write "<th>Jamb</th><th>Head</th><th>Int</th><th>Sill</th><th>Strut</th><th>Door Adapter</th><th>Doors</th></tr>" 
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Job") & "</td><td>" & RS("Floor") &"</td><td>" & RS("datestamp") & "</td><td>" & RS("cycles") & "</td><td>" & RS("Totalsqft") & " ft<sup>2</sup></td><td>" & RS("TotalWin") & "</td>"
	response.write "<td>" & RS("ftawnings") & " ft / # " & RS("totalawnings") & "</td>"
	response.write "<td>" & RS("sqftpanels") & " ft<sup>2</sup> / # " & RS("totalpanels") & "</td>"
	response.write "<td>" & RS("sqftSP") & " ft<sup>2</sup> / # " & RS("totalSP") & "</td>"
	response.write "<td>" & RS("sqftSU_SMALL") & " ft<sup>2</sup> / # " & RS("totalSU_SMALL") & "</td>"
	response.write "<td>" & RS("sqftSU_MEDIUM") & " ft<sup>2</sup> / # " & RS("totalSU_MEDIUM") & "</td>"
	response.write "<td>" & RS("sqftSU_LARGE") & " ft<sup>2</sup> / # " & RS("totalSU_LARGE") & "</td>"
	response.write "<td>" & RS("ftJamb") & " ft</td><td>" & RS("ftHead") &" ft</td><td>" & RS("ftInt") & " ft</td><td>" & RS("ftSill") & " ft</td><td>" & RS("ftstrut") & " ft</td><td>" & RS("ftdooradapter") & " ft</td><td>" & RS("totaldoor") & "</td>"
	
	response.write "</tr>"
	
rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing


%>
               
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
