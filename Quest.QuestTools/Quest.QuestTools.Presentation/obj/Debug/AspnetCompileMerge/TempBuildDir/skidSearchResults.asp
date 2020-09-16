<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Search Results - Show Search Results for SkidSearch -->
<!-- Shows Results for Barcode, Job, Floor, Tag-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Skid Results</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
 </head>

<body>
     
     

     


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="SkidSearch.asp" target="_self">Search Again</a>
		<a class="button" type="cancel" href="ScanHome.html#_Skids" target="_self">Skids </a>
        </div>
  <ul id="Profiles" title="Skids and Skid Items" selected="true">
  
  
   <% 

currentDate = Date()



bc = UCASE(request.querystring("barcodeid"))
job = UCASE(request.querystring("job"))
floor = UCASE(request.querystring("floor"))
tag = UCASE(request.querystring("tag"))


if bc = "" AND job = "" AND floor = "" AND tag = "" then

else

'	if bc = "" then 
'		bc = "%"
'	end if
'	if job = "" then 
'		bc = "%"
'	end if
'	if floor = "" then 
'		floor = "%"
'	end if
'	if tag = "" then 
'		tag = "%"
'	end if

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM SKIDITEM WHERE barcode like '%" & bc & "%' AND job like '%" & job & "%' AND floor like '%" & floor & "%' AND tag like '%" & tag & "%' ORDER BY ID ASC"
	'strSQL = "SELECT * FROM SKIDITEM WHERE barcode like '%%' ORDER BY ID ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	
	
			response.write "<li>Skid Items Matching Query</li>"
			response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='10%'>SKID</th><th width='30%'>Barcode</th><th width='10%'>Job</th><th width='10%'>Floor</th><th width='10%'>Tag</th><th  width='20%'>Scan Date</th><th  width='10%'>Flushed</th></tr>"
	
		if not rs.bof then
		
			rs.movefirst
			Do while not rs.eof
				response.write "<tr><td>" & rs("name") & "</td><td>" & rs("Barcode") & "</td><td>" & rs("Job") & "</td><td>" & rs("Floor") & "</td><td>" & rs("Tag") & "</td><td>" & rs("ScanDate") & "</td><td>" & rs("FlushedDate") & "</td></tr>"
			rs.movenext
			loop
		else
			response.write "<tr ><td colspan = '7'>Empty Search Results: Please try again</td></tr>"
		end if
	response.write "</table></li>"
	rs.close
	set rs = nothing
	
end if
DBConnection.close
set DBConnection = nothing



 %>
     </ul>


</body>
</html>