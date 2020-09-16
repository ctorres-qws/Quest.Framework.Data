<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->
<!-- MODIFIED BY Annabel Ramirez Feb 12, 2020 - to pull values for IntDoorType and ExtDoorType -->

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

cid = REQUEST.QueryString("CID")

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="DoorStyle.asp" target="_self">Door Style</a>
    </div>
    
      <%                  

gi_Mode = c_MODE_SQL_SERVER
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM StylesDoor"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & cid
%>
    
    
        <form id="edit" title="Edit Door Style" class="panel" name="edit" action="DoorStyleEditConf.asp" method="GET" target="_self" selected="true" > 
        <h2>Edit Door Style</h2>
  
        <fieldset>


        <div class="row">
            <label>Name</label>
            <input name='Name' type="text" id='Name' value="<% response.write rs.fields("Name") %>" >
        </div>
		
		<div class="row">
            <label>Job</label>
            <input name='Job' type="text" id='Job' value="<% response.write rs.fields("Job") %>" >
		
		<div class="row">
			<label>Int</label>
			<Select name='IntDoorType'>
				<option value="Fapim" <% if Trim(rs.fields("IntDoorType")) = "Fapim" then response.write "Selected"%> >Fapim</option>
				<option value="Metra" <% if Trim(rs.fields("IntDoorType")) = "Metra" then response.write "Selected"%> >Metra</option>
			</Select>				
		</div>			

		<div class="row">
			<label>Ext</label>
			<Select name='ExtDoorType'>
				<option value="Fapim" <% if Trim(rs.fields("ExtDoorType")) = "Fapim" then response.write "Selected"%> >Fapim</option>
				<option value="Hopi" <% if Trim(rs.fields("ExtDoorType")) = "Hopi" then response.write "Selected"%> >Hopi</option>
				<option value="Metra" <% if Trim(rs.fields("ExtDoorType")) = "Metra" then response.write "Selected"%> >Metra</option>
				<option value="None" <% if Trim(rs.fields("ExtDoorType")) = "None" then response.write "Selected"%> >None</option>
			</Select>				
		</div>		
		
        <input type="hidden" name='cid' id='cid' value="<%response.write rs.fields("id") %>" >
                      
</fieldset>
        <br>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
</form> 
  
<% 
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
End Function	
%>

          
    
</body>
</html>