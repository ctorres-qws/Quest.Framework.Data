
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Styles</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job</a>
    </div>
    
        
    
              <form id="edit" title="Select Panel Style by Job" class="panel" name="edit" action="PanelStylebyJob.asp" method="GET" target="_self" selected="true" > 
        <h2>Panel Styles by Job</h2>
  

<fieldset>
     <div class="row">
<label>Parent Job </label>
<input type="text" name = "Parent" id="Parent" />
            </div>
            
</fieldset>

        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Search ALL Panel Styles by Parent Job </a>
		<a class="whiteButton" href="PanelStyleEnter.asp">Add New Style </a><BR><BR><BR><BR>
		<ul id="Profiles" title="Panel Styles - Job Search" selected="true">
	<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM StylesPanel ORDER BY Parent, NAME ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection		

response.write "<li class='group'>All Panel Styles</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Name</th><th>Description</th><th>Parent</th><th>Color Code</th><th>Side</th><th>Material</th><th>Colour</th><th>Notes</th><th>Edit</th></tr>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Name") & "</td>"
	response.write "<td>" & RS("Description") & "</td>"
	response.write "<td>" & RS("PARENT") &"</td>"
	response.write "<td>" & RS("COLORCODE") &"</td>"
	response.write "<td>" & RS("Side") & "</td>"
	response.write "<td>" & RS("Material") & "</td>"
	response.write "<td>" & RS("Colour") & "</td>"
	response.write "<td>" & RS("Notes") & "</td>"
	response.write "<td><a href =><a href='PanelStyleEditForm.asp?cid=" & rs("ID") & "' target='_self' >Manage</td>"
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing
%>		
		
		 </form> 

</body>
</html>


