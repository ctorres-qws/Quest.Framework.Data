                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>JPanel Colours</title>
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
	
	ChildColour = False
	Child = Request.querystring("Parent")
	Parent = Request.querystring("Parent")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Job, Parent FROM Z_Jobs WHERE [Job] = '" & Parent & "' ORDER BY Job ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

if RS("JOB") <> rs("Parent") then
	Parent = rs("Parent")
	ChildColour= True
end if

rs.close
set res = nothing

	
	
Set rs = Server.CreateObject("adodb.recordset")
if UCASE(PARENT) = "ALL" then
	strSQL = "SELECT * FROM StylesPanel ORDER BY Parent, NAME ASC"
else
	strSQL = "SELECT * FROM StylesPanel WHERE [Parent] = '" & Parent & "' ORDER BY NAME ASC"
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
       <a class="button leftButton" type="cancel" href="PanelStylebyJob1.asp" target="_self">New Search</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Panel Styles - Job Search" selected="true">
<% 
response.write "<li class='group'>All Panel Styles for Parent Color - " & Parent & " </li>"
response.write "<li>" & Child & " is a Child Colour of " & Parent & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Name</th><th>Description</th><th>Parent</th><th>Color Code</th><th>Side</th><th>Material</th><th>Colour</th><th>Notes</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("Name") & "</td><td>" & RS("Description") & "</td><td>" & RS("PARENT") &"</td><td>" & RS("COLORCODE") &"</td><td>" & RS("Side") & "</td><td>" & RS("MAterial") & "</td><td>" & RS("Colour") & "</td><td>" & RS("NOTES") & "</td>"
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
