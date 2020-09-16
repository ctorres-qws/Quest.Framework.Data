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
	

	DIM JOBFLOOR()
	DIM JOB()
	DIM FLOOR()
	Arraysize = Request.QueryString("JOBFLOOR").Count
	ReDim JOBFLOOR(Arraysize)
	ReDim JOB(Arraysize)
	ReDim FLOOR(Arraysize)
	
for i=1 to Request.QueryString("JOBFLOOR").Count
CJF = Request.QueryString("JOBFLOOR")(i)
	JOB(i) = Left(CJF, 3)
	if LEN(CJF)= 5 then
		FLOOR(i) = RIGHT(CJF, 1)
	end if
	if LEN(CJF)= 6 then
		FLOOR(i) = RIGHT(CJF, 2)
	end if
	if LEN(CJF) = 7 then
		FLOOR(i) = RIGHT(CJF, 3)
	end if
next


for i=1 to Request.QueryString("JOBFLOOR").Count
SearchLine = SearchLINE & " OR (JOB = '" & JOB(i) 
SearchLine = SearchLINE & "' AND FLOOR = '" & FLOOR(i) & "')"
next
if arraysize>0 then
SEARCHLINE = RIGHT( SEARCHLINE, LEN(SEARCHLINE)-4)
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM PickList WHERE " & SEARCHLINE & " ORDER BY JOB, FLOOR, DIE"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
end if


%>	
	
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="PicklistSearch.asp" target="_self">PL Search</a>
        </div>
        
        <ul id="screen1" title="View PL Records" selected="true">
    <%


response.write "<li class='group'>Pick List Filtered by " & SEARCHTYPE & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Die</th><th>Colour</th><th>Length</th><th>QTY</th><th>Pick Date</th><th>Entry Date</th></tr>"
if arraysize>0 then
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
	response.write "<td>" & RS("PICKDATE") & "</td>"
	response.write "<td>" & RS("ENTRYDATE") & "</td>"
	response.write "</tr>"
	
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing

else
Response.write "<tr><td>No Search Items</td></tr>"
end if
DBConnection.close
set DBConnection = nothing

%>
      </ul>                 
               
               
</body>
</html>
