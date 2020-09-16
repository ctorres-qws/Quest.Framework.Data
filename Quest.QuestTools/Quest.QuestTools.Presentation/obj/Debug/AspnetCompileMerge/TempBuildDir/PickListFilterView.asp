<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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
strSQL = "SELECT * FROM PickList ORDER BY JOB, COLOUR"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
	
	

JOB = request.QueryString("Job")	
FLOOR = request.QueryString("Floor")
COLOUR = request.QueryString("Colour")	
DIE = request.QueryString("Die")	
PickDate = request.QueryString("PickDate")	

If JOB <> "" and JOB <> "ANY" then
		rs.filter = "JOB = '" & JOB & "'"
		FilterCodes = FilterCodes & " - " & JOB
end if
If FLOOR <> "" and FLOOR <> "ANY" then
		rs.filter = "FLOOR = '" & FLOOR & "'"
		FilterCodes = FilterCodes & " - " & FLOOR
end if
If COLOUR <> "" and COLOUR <> "ANY" then
		rs.filter = "COLOUR = '" & COLOUR & "'"
		FilterCodes = FilterCodes & " - " & COLOUR
end if
If DIE <> "" and DIE <> "ANY" then
		rs.filter = "DIE = '" & DIE & "'"
		FilterCodes = FilterCodes & " - " & DIE
end if

If ISDate(PickDate) = True then
		rs.filter = "PICKDATE > #" & DATE() & "# AND PICKDATE < #" & PickDate & "#" 
		FilterCodes = FilterCodes & " - Between " & DATE & " And " & PickDate
end if


%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="PickListFIlter.asp" target="_self">Filter</a>
        </div>
        
        <ul id="screen1" title="View PL Records" selected="true">
		<li> Filter by <% response.write FilterCodes %> </li>
    <%


response.write "<li class='group'>Pick List Filtered by " & SEARCHTYPE & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Die</th><th>Colour</th><th>Length</th><th>QTY</th><th>Pick Date</th><th>EntryDate</th></tr>"


if rs.eof then
Response.write "<tr><td>No current Items</td></tr>"
end if		
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JOB") & "</td>"
	response.write "<td>" & RS("FLOOR") & "</td>"
	response.write "<td>" & RS("DIE") & "</td>"
	response.write "<td>" & RS("COLOUR") & "</td>"
	response.write "<td>" & RS("LENGTH") & "</td>"
	response.write "<td>" & RS("QTY") & "</td>"
	response.write "<td>" & RS("PickDate") & "</td>"
	response.write "<td>" & RS("ENTRYDATE") & "</td>"
	response.write "</tr>"
	
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
