                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Shear Test Program sitting on Blue Zipper Machine stores information on Shear test-->
<!-- Currently a Manual test with future interest in forcing per material rolled -->
<!-- created for LEv Bedoev and Danial Zalcman -->
<!-- Created April 22nd, 2015 by Michael Bernholtz -->

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
strSQL = "SELECT * FROM PROZipperShearTest ORDER BY ID DESC"
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
        <a class="button leftButton" type="cancel" href="index.html#_QCT" target="_self">QC</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Zipper Shear Test Results" selected="true">
        
        
<% 
response.write "<li class='group'>Shear Test Results </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li>Jan 2020 - Red has been renamed to Blue Zipper 2</li>"
response.write "<li>A Fail is recorded if either Ext. Frame or Int. Frame falls below 200 KPA</li>"
response.write "<li>A Warning is recorded if either Ext. Frame or Int. Frame falls Below 220 KPA but is Above 200 KPA</li>"
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Date</th><th>JOB</th><th>Zipper</th><th>Ext Frame(KPA)</th><th>Int Frame(KPA)</th><th>Employee</th><th>Notes</th><th>Pass/Fail</th></tr>"
do while not rs.eof
	passfail = "PASS"
	if RS("FrameEXT") < 220  or RS("FrameINT") < 220 then
		if RS("FrameEXT") < 200  or RS("FrameINT") < 200 then
			passfail = "FAIL"
		else 
			passfail = "Warning"
		end if
	end if
	
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("DATETIME") & "</td>"
	response.write "<td>" & RS("JOB") &"</td><td>" & RS("ZIPPERName") & "</td>"
	response.write "<td>" & RS("FrameExt") & " KPA</td><td>" & RS("FRAMEINT") & " KPA</td>"
	response.write "<td>" & RS("Employee") & "</td><td>" & RS("Note") & "</td><td>" & passfail & "</td>" 

	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li></ul>"



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
