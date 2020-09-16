<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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

 
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PanelStyleByJob1.asp?id=<% response.write pid %>" target="_self">Edit Style</a>
    </div>
    
      
    
<form id="conf" title="Edit Panel Style" class="panel" name="conf" action="PanelStyleByJob1.asp#_screen1" method="GET" target="_self" selected="true" >              

  
   
        
  
<%    
'Added extra code to match the Job table to the Colour Table, January 2015, Michael Bernholtz
              
cid = request.querystring("cid")

NAME = REQUEST.QueryString("Name")
DESCRIPTION = REQUEST.QueryString("Description")
PARENT = REQUEST.QueryString("Parent")
COLORCODE = REQUEST.QueryString("ColorCode")
SIDE = REQUEST.QueryString("SIDE")
MATERIAL = REQUEST.QueryString("Material")
COLOUR = REQUEST.QueryString("Colour")
NOTES = REQUEST.QueryString("Notes")

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

			'Set Color Update Statement
				StrSQL = FixSQLCheck("UPDATE StylesPanel  SET PARENT= '"& PARENT & "', COLORCODE = '"& COLORCODE & "', NAME = '"& NAME & "', DESCRIPTION = '"& DESCRIPTION & "', SIDE = '"& SIDE & "', MATERIAL = '" & MATERIAL & "', COLOUR ='" & COLOUR & "', NOTES= '" & NOTES & "' WHERE ID = " & CID, isSQLServer)
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)

DbCloseAll

End Function

%>

<h2>Color: <%response.write Project %> Edited</h2>
        <BR>
	<ul>
    <li> Panel Style Edited:</li>
	<li><% response.write "Parent Colour: " & PARENT %></li>
	<li><% response.write "Colour Code: " & COLORCODE %></li>
	<li><% response.write "NAME: " & Name %></li>
	<li><% response.write "Description: " & DESCRIPTION %></li>
	<li><% response.write "Side: " & SIDE %></li>
    <li><% response.write "Material: " & MATERIAL %></li>
	<li><% response.write "Colour: " & COLOUR %></li>
	<li><% response.write "Notes: " & NOTES %></li>
		

	</ul>
         <a class="whiteButton" href="javascript:conf.submit()">Panel Styles</a>
            
            </form>

  
<% 



'DBConnection.close
'set DBConnection=nothing
%>

          
    
</body>
</html>
