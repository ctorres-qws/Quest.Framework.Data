<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--AllJobs Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->
<!-- Updated for Global Variables Updates by Annabel Ramirez August 2019 -->

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
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="AllJobsChildENTER.asp" target="_self">Enter Job</a>

    </div>

<ul id="Report" title="Added" selected="true">

<%
Passkey = "QUEJAY"
Password = UCASE(TRIM(Request.Form("pwd")))

JOB = TRIM(UCASE(REQUEST.QueryString("JOB")))
PARENT = UCASE(REQUEST.QueryString("PARENT"))

'SILL = REQUEST.QueryString("SILL")
GLASSFLUSH = REQUEST.QueryString("GLASSFLUSH")
BEAUTYSTYLE = REQUEST.QueryString("BEAUTYSTYLE")
COLORMATCH = REQUEST.QueryString("COLORMATCH")
PANELPUNCH = REQUEST.QueryString("PANELPUNCH")
R3VentSize = REQUEST.QueryString("R3VENTSIZE")
EXT_COLOUR = REQUEST.QueryString("EXT_COLOUR")
INT_COLOUR = REQUEST.QueryString("INT_COLOUR")
FLOORS = REQUEST.QueryString("FLOORS")
MATERIAL = REQUEST.QueryString("MATERIAL")

JOB_NAME = REQUEST.QueryString("JOB_NAME")
JOB_ADDRESS = REQUEST.QueryString("JOB_ADDRESS")
JOB_CITY = REQUEST.QueryString("JOB_CITY")
JOB_COUNTRY = REQUEST.QueryString("JOB_COUNTRY")
MANAGER = REQUEST.QueryString("MANAGER")
RLIST = REQUEST.QueryString("RLIST")
MATERIAL = REQUEST.QueryString("MATERIAL")
COMPLETED = REQUEST.QueryString("COMPLETED")
SDSILL = REQUEST.QueryString("SDsill")
SWSILL = REQUEST.QueryString("SWsill")
%>
<%
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

If Password = Passkey Then
	DBOpen DBConnection, isSQLServer
	Exists = False
	ErrorComment = ""

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_Jobs WHERE PARENT = '" & Parent & "' Order by JOB ASC"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection
	RS.filter = "JOB = '" & JOB & "'"
		if (rs.eof and Len(JOB)=3) then
		else 
			if Len(Job) = 3 then
				Exists = True
				ErrorComment = "Already Exists in the Database, Please Edit the existing record"
			else
				Exists = True
				ErrorComment = "Please enter a 3 Letter Job Code"
			End if
		end if
	
	RS.filter = "JOB = '" & Parent & "'"

JOB_NAME = RS("Job_Name")
JOB_ADDRESS = RS("Job_Address")
JOB_CITY = RS("Job_City")
JOB_COUNTRY = RS("Job_Country")
MATERIAL = RS("Material")
MANAGER = RS("Manager")
MANAGEREMAIL = RS("ManagerEmail")
ENGINEER = RS("Engineer")
ENGINEEREMAIL = RS("EngineerEmail")
' FLOORS = RS("Floors")
RLIST = RS("RecipientList")
IMPREC = RS("ImporterRecord")
IMPADD = RS("ImporterAddress")
IMPTAX = RS("ImporterTaxID")
EXPTAX = RS("ExporterTaxID")
COMPLETED = RS("COMPLETED")
ONSITEDATE=RS("OnSiteDate")
'If COMPLETED = "on" then
'	COMPLETED = TRUE
'Else
'	COMPLETED = FALSE
'End If
'If FLOORS = "" then
'	FLOORS = 0
'End If
	
	
'SILL = REQUEST.QueryString("sill")
GLASSFLUSH = REQUEST.QueryString("GLASSFLUSH")
BEAUTYSTYLE = REQUEST.QueryString("BEAUTYSTYLE")
COLORMATCH = REQUEST.QueryString("COLORMATCH")
PANELPUNCH = REQUEST.QueryString("PANELPUNCH")
R3VentSize = REQUEST.QueryString("R3VENTSIZE")
EXT_COLOUR = REQUEST.QueryString("EXT_COLOUR")
INT_COLOUR = REQUEST.QueryString("INT_COLOUR")
FLOORS = REQUEST.QueryString("FLOORS")
SDSILL = REQUEST.QueryString("SDsill")
SWSILL = REQUEST.QueryString("SWsill")

If FLOORS = "" then
	FLOORS = 0
End If
If R3VentSize = "" then
	R3VentSize = 0
End If
If Exists = False then
	rs.AddNew
	rs.Fields("JOB") = JOB
	rs.Fields("PARENT") = PARENT
	rs.Fields("JOB_NAME") = JOB_NAME
	rs.Fields("JOB_ADDRESS") = JOB_ADDRESS
	rs.Fields("JOB_COUNTRY") = JOB_COUNTRY
	rs.Fields("JOB_CITY") = JOB_CITY
	rs.Fields("MATERIAL") = MATERIAL
	rs.Fields("FLOORS") = INT(FLOORS)
	rs.Fields("RecipientList") = RLIST
	rs.Fields("ImporterRecord") = IMPREC
	rs.Fields("ImporterAddress") = IMPADD
	rs.Fields("ImporterTaxID") = IMPTAX
	rs.Fields("ExporterTaxID") = EXPTAX
	rs.Fields("EXT_COLOUR") = EXT_COLOUR
	rs.Fields("INT_COLOUR") = INT_COLOUR
	'rs.Fields("SILL") = SILL
	rs.Fields("COLORMATCH") = COLORMATCH
	rs.Fields("GLASSFLUSH") = GLASSFLUSH
	rs.Fields("BEAUTYSTYLE") = BEAUTYSTYLE
	rs.Fields("COMPLETED") = COMPLETED
	rs.Fields("MANAGER") = MANAGER
	rs.Fields("MANAGEREMAIL") = MANAGEREMAIL
	rs.Fields("ENGINEER") = ENGINEER
	rs.Fields("ENGINEEREMAIL") = ENGINEEREMAIL
	rs.Fields("PANELPUNCH") = PANELPUNCH
	rs.Fields("R3VENTSIZE") = INT(R3VENTSIZE)
	rs.Fields("OnSiteDate") = ONSITEDATE
	rs.Fields("SDSILL") = SDSILL
	rs.Fields("SWSILL") = SWSILL

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	rs.update

	Call StoreID1(isSQLServer, rs.Fields("ID"))

	end if

	

	if Exists = True then
	%>
		<li><% response.write "JOB: " & JOB %></li>
		<li><% response.write "Already Exists in the Database, Please Edit the existing record" %></li>
<%
	else
%>

	<li><% response.write "JOB: " & JOB %></li>
	<li><% response.write "PARENT: " & PARENT %></li>
	<li><% response.write "FULL NAME: " & JOB_NAME %></li>
	<li><% response.write "ADDRESS: " & JOB_ADDRESS %></li>
	<li><% response.write "CITY: " & JOB_CITY %></li>
	<li><% response.write "MATERIAL: " & MATERIAL %></li>
	<li><% response.write "MANAGER: " & MANAGER %></li>
	<li><% response.write "FLOORS: " & FLOORS %></li>
	<li><% response.write "Recipient List: " & RLIST %></li>
	<li><% response.write "EXTERIOR COLOUR: " & EXT_COLOUR %></li>
	<li><% response.write "INTERIOR COLOUR: " & INT_COLOUR %></li>
	<li><% response.write "HBAR Color Match: " & COLORMATCH %></li>
	<li><% response.write "COMPLETED: " & COMPLETED %></li>
	
	
	<%
	end if
	%>
	<li><center><a href="alljobsreport.asp"> Back to All Jobs List </a></center></li>
</ul>

<%

	DbCloseAll
End If
End Function

'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection = nothing
%>
<% 
if Password = Passkey then
%> 
<form id="conf" title="Child Added" class="panel" name="conf" action="index.html#_Job" method="GET" target="_self" selected="true" >              
	
        <h2>Child Addition</h2>       
<%
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="AllJobsChildConf.asp?JOB=<%response.write JOB%>&PARENT=<%response.write PARENT%>&JOB_NAME=<%response.write JOB_NAME%>&JOB_ADDRESS=<%response.write JOB_ADDRESS%>&JOB_CITY=<%response.write JOB_CITY%>&JOB_COUNTRY=<%response.write JOB_COUNTRY%>&MATERIAL=<%response.write MATERIAL%>&MANAGER=<%response.write MANAGER%>&ENGINEER=<%response.write ENGINEER%>&ENGINEEREMAIL=<%response.write ENGINEEREMAIL%>&RLIST=<%response.write RLIST%>&IMPREC=<%response.write IMPREC%>&IMPADD=<%response.write IMPADD%>&IMPTAX=<%response.write IMPTAX%>&EXPTAX=<%response.write EXPTAX%>&COMPLETED=<%response.write COMPLETED%>&GLASSFLUSH=<%response.write GLASSFLUSH%>&BEAUTYSTYLE=<%response.write BEAUTYSTYLE%>&COLORMATCH=<%response.write COLORMATCH%>&PANELPUNCH=<%response.write PANELPUNCH%>&R3VentSize=<%response.write R3VentSize%>&EXT_COLOUR=<%response.write EXT_COLOUR%>&INT_COLOUR=<%response.write INT_COLOUR%>&FLOORS=<%response.write FLOORS%>&SDSILL=<%response.write SDSILL%>&SWSILL=<%response.write SWSILL%>" method="post" target="_self" selected="True">
	<ul id="Report" title="Password" selected="true">
<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
end if
if Password = Passkey then
%>
	<ul id="Report" title="Added" selected="true">

<%	
		Response.Write "<li> JOB: " & JOB & "</li>"
		Response.Write "<li> PARENT: " & PARENT & "</li>"
		Response.Write "<li> FULL NAME: " & JOB_NAME & "</li>"
		Response.Write "<li> ADDRESS: " & JOB_ADDRESS & "</li>"
		Response.Write "<li> CITY: " & JOB_CITY & "</li>"
		Response.Write "<li> MATERIAL: " & MATERIAL & "</li>"
		Response.Write "<li> MANAGER: " & MANAGER & "</li>"
		Response.Write "<li> FLOORS: " & FLOORS & "</li>"
		Response.Write "<li> Recipient List: " & RLIST & "</li>"
		Response.Write "<li> EXTERIOR COLOUR: " & EXT_COLOUR & "</li>"
		Response.Write "<li> INTERIOR COLOUR: " & INT_COLOUR & "</li>"
		Response.Write "<li> HBAR Color Match: " & COLORMATCH & "</li>"		
		Response.Write "<li> COMPLETED: " & COMPLETED & "</li>"		
%>	
        <BR>
         <a class="whiteButton" href="AllJobsReport.asp#_Profiles" target="_self"> Back</a>
            </form>	
<%	
end if
%>
     

</body>
</html>

