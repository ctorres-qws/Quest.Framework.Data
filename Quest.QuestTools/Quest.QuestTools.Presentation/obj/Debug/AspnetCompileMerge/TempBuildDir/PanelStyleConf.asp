<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Updated to include new Entry items and update them to the Database by Michael Bernholtz on request of Jody Cash-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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

Parent = Request.QueryString("Parent")
ColorCode = Request.QueryString("ColorCode")
DESCRIPTION = Request.QueryString("DESCRIPTION")
NAME = Request.QueryString("NAME")
NOTES = Request.QueryString("NOTES")
SIDE = Request.QueryString("SIDE")
COLOUR = Request.QueryString("COLOUR")
MATERIAL = Request.QueryString("MATERIAL")

Set rs = Server.CreateObject("adodb.recordset")
	
	strSQL = "SELECT * FROM StylesPanel WHERE ID=-1"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection

	rs.AddNew
	rs.Fields("Name") = Name
	rs.Fields("Description") = Description
	rs.Fields("Parent") = Parent
	rs.Fields("ColorCode") = ColorCode
	rs.Fields("Notes") = Notes
	rs.Fields("Side") = Side
	rs.Fields("Material") = Material
	rs.Fields("Colour") = Colour
	rs.update
	rs.close
	set rs = nothing

	
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PanelStyleEnter.asp" target="_self">Panel Entry</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>


    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Name " & Name %></li>
	<li><% response.write "Description " & Description %></li>
    <li><% response.write "Parent " & parent %></li>
	<li><% response.write "Color Code " & ColorCode %></li>
	<li><% response.write "Side " & Side %></li>
    <li><% response.write "Material " & Material %></li>
    <li><% response.write "Colour " & Colour %></li>

	<li><% response.write "Notes " & Notes %></li>

	</ul>

<% 



DBConnection.close
set DBConnection = nothing
%>

</body>
</html>



