<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- June 2016, for Shaun Levy, Created by Michael Bernholtz-->
<!-- Individual Report to describe the Number of windows and Average Number of Openings in a Job -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Average Openings per Job</title>
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
strSQL = "SELECT * FROM Z_JOBS  ORDER BY JOB ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Openings per Job</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Openings Per Job" selected="true">
        
        
<% 
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>JOB</th><th># of Windows</th><th>Total Openings</th><th>Average Openings </th></tr>"
do while not rs.eof

Set rsJOB = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [" & rs("JOB") & "] ORDER BY JOB ASC"
	On Error Resume Next  
		rsJOB.Open strSQL, DBConnection
	On Error GoTo 0
	
	If rsJOB.State = 1 Then 
	
	TotalOpenings = 0
	TotalWindows = 0
	TotalOther = 0
	AverageOpenings = 0
	Do while not rsJOB.eof
		if LEN(RSJOB("STYLE")) = 4 then
			TotalWindows = TotalWindows + 1
			WinOpens = Left(RSJOB("STYLE"),1) + 0
			TotalOpenings = TotalOpenings + WinOpens
		else 
			TotalOther = TotalOther+1
		End if
	
	rsJOB.movenext
	loop
	
		response.write "<tr>"
		response.write "<td>" & RS("JOB") & "</td><td>" & TotalWindows & "</td><td>" & TotalOpenings &"</td><td>" & TotalOpenings/TotalWindows & "</td>"
		response.write "</tr>"
	
		rsJOB.Close
		
	End If
	set rsJOB = nothing
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
