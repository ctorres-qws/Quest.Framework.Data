<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- CUT LIST View list designed by Michael Bernholtz -->
<!-- Slava Kotok and Jody Cash requested a view of all the Cutlists in the system -->
<!-- THis is a simple viewing List of those Cutlists - Showing Cutlist and associated Details -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>CutList View</title>
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

Server.ScriptTimeout=200
	
FilterJob = REQUEST.QUERYSTRING("FILTERJOB")	
FilterFloor = REQUEST.QUERYSTRING("FILTERFLOOR")

cutPiece = 0
TotalPiece = 0

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Cut List View</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
   
			<ul id="Profiles" title="All - Cut List View " selected="true">
			<form id="Cutlist" action="CutListView.asp" name="Cutlist"  method="GET" target="_self">  
			<h2>Filter List by Job, Floor:</h2>
			<fieldset>
				<div class="row">
					<label>Job</label>
					<input type="text" name='FilterJob' id='FilterJob' value = '<% response.write FilterJob %>' />
				</div>
				<div class="row">
					<label>Floor</label>
					<input type="text" name='FilterFloor' id='FilterFloor'value = '<% response.write FilterFloor %>' />
				</div>
				<a class="whiteButton" onClick="Cutlist.submit()">Filter</a><BR>
			</fieldset>
			</form>
									
<%

If UCase(Request("Search")) <> "NEW" Then

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Top 200 * FROM Z_CUTLISTS ORDER BY ID DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

if FilterJob = "" then
else
	if FilterFloor = "" then
		rs.filter = "JOB = '" & FilterJob & "'"
	else
		rs.filter = "JOB = '" & FilterJob & "'AND  Floor LIKE '%" & FilterFloor & "%'"
	end if
end if
 
response.write "<li class='group'>Cut List View</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Cutlist</th><th>Job</th><th>Floor</th><th>Cycle</th><th>Processed Date</th><th>Cut</th><th>Total</th></tr>"
do while not rs.eof


	response.write "<tr>"
	response.write "<td>" & RS("CutList") & "</td>"
	response.write "<td>" & RS("Job") &"</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("Cycle") & "</td>"
	response.write "<td>" & RS("Day") & "/" & RS("Month") & "/" & RS("Year") & "</td>"
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT cStatus FROM " & RS("CutList")
	On Error Resume Next 
	rs2.Open strSQL2, DBConnection
	On Error GoTo 0
	TotalData = 0
	cutData = 0
	If rs2.State = 1 Then 
	do while not rs2.eof 
		if rs2("cStatus") = -1 then
				CutData = CutData + 1
		end if
				TotalData = TotalData + 1
	rs2.movenext
	loop
	rs2.close
	set rs2 = nothing
	end if
	
	if TotalData = 0 then
	Response.write "<td></td><td></td>"
	else
	Response.write "<td>" & cutdata & "</td><td>" & TotalData &"</td>"
	end if
	response.write "</tr>"
rs.movenext
loop
	


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

End If

%>
      </table></li>         
    </ul>        
     		       
            
       
            
              
               
                
             
               
</body>
</html>
