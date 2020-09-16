                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Window Production - Job Floor representation of Window Number and Total SQFT-->
<!-- Collecting from X_WIN_ARCHIVE2 which already stores all of this data-->
<!-- Created July 31st, 2017 by Michael Bernholtz - For Jody Cash and Shaun Levy-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Service Glass Report</title>
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
strSQL = "SELECT DATESTAMP, JOB, FLOOR, TOTALWIN, TOTALPanels, TOTALDoor, TOTALSQFT FROM X_WIN_ARCHIVE2 WHERE [DATESTAMP]< #09/01/2014# AND [DATESTAMP]> #08/31/2013# ORDER BY JOB ASC, FLOOR DESC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Job / Floor - SQFT" selected="true">
        
        
<% 
TotalSQFT = 0
response.write "<li class='group'>JOB / FLOOR - SQFT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li> ALL Window Production info from August 31st, 2013 to September 1st, 2014</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>DATE</th><th>Job</th><th>Floor</th><th>Total Windows</th><th>Total Panels</th><th>Total Doors</th><th>Total SQFT</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("DATESTAMP") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TOTALWIN") & "</td><td>" & RS("TOTALPanels") & "</td><td>" & RS("TOTALDoor") & "</td><td>" & RS("TOTALSQFT") & "</td>"
	response.write " </tr>"
	TotalSQFT = TOTALSQFT + RS("TOTALSQFT")
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

response.write "<li>TOTAL SQFT: " & TOTALSQFT & "</li>"
%>
               
<li>//END//</li>
/
      </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
