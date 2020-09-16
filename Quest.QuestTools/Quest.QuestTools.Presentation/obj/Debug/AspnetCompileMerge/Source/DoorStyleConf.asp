<!--#include file="dbpath.asp"-->                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->

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

NAME = REQUEST.QueryString("Name")
JOB = REQUEST.QueryString("Job")
EXTDOORTYPE = REQUEST.QueryString("ExtDoorType")
INTDOORTYPE  = REQUEST.QueryString("IntDoorType")

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
	
		Dim PROGRAMIN: PROGRAMIN = ""
		Dim PROGRAMOUT: PROGRAMOUT = ""

		If INTDOORTYPE = "Fapim" Then
			If EXTDOORTYPE = "Fapim" Then
				PROGRAMIN = "SW1"
				PROGRAMOUT = "SW6"
			ElseIf EXTDOORTYPE = "Hopi" Then
				PROGRAMIN = "SW2"
				PROGRAMOUT = "SW7"
			ElseIf EXTDOORTYPE = "None" Then
				PROGRAMIN = "SW3"			
				PROGRAMOUT = "SW8"
			End If
		ElseIf INTDOORTYPE = "Metra" Then
			If EXTDOORTYPE = "Metra" Then
				PROGRAMIN = "SW4"
				PROGRAMOUT = "SW9"
			ElseIf EXTDOORTYPE = "Hopi" Then
				PROGRAMIN = "SW5"
				PROGRAMOUT = "SW10"
			End If
		End If

Set rs = Server.CreateObject("adodb.recordset")
	
	strSQL = "SELECT * FROM StylesDoor WHERE ID=-1"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection

	rs.AddNew
	rs.Fields("Name") = NAME
	rs.Fields("Job") = JOB
	rs.Fields("ExtDoorType") = EXTDOORTYPE
	rs.Fields("IntDoorType") = INTDOORTYPE
	rs.Fields("ProgramIn") = PROGRAMIN
	rs.Fields("ProgramOut") = PROGRAMOUT
	rs.update
	rs.close
	set rs = nothing
DBConnection.close
set DBConnection = nothing
End Function	

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="DoorStyleEnter.asp" target="_self">Door Entry</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Name: " & NAME %></li>
	<li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Interior Door Type: " & INTDOORTYPE %></li>	
    <li><% response.write "Exterior Door Type: " & EXTDOORTYPE %></li>
	</ul>

</body>
</html>



