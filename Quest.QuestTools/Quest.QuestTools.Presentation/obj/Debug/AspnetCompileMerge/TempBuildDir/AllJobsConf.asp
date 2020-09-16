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

<%
Passkey = "QUEJAY"
Password = UCASE(TRIM(Request.Form("pwd")))


Exists = False
ErrorComment = ""				 

JOB = UCASE(REQUEST.QueryString("JOB"))
PARENT = UCASE(REQUEST.QueryString("PARENT"))
JOB_NAME = REQUEST.QueryString("JOB_NAME")

JOB_ADDRESS = REQUEST.QueryString("JOB_ADDRESS")
JOB_CITY = REQUEST.QueryString("JOB_CITY")
JOB_COUNTRY = REQUEST.QueryString("JOB_COUNTRY")
MATERIAL = REQUEST.QueryString("MATERIAL")
'SILL = REQUEST.QueryString("sill")
GLASSFLUSH = REQUEST.QueryString("GLASSFLUSH")
BEAUTYSTYLE = REQUEST.QueryString("BEAUTYSTYLE")
MANAGER = REQUEST.QueryString("Manager")
MANAGEREMAIL = REQUEST.QueryString("ManagerEmail")
ENGINEER = REQUEST.QueryString("ENGINEER")
ENGINEEREMAIL = REQUEST.QueryString("ENGINEEREMAIL")
FLOORS = REQUEST.QueryString("FLOORS")
COLORMATCH = REQUEST.QueryString("COLORMATCH")

EXTGLASS = REQUEST.QueryString("EXTGLASS")
INTGLASS = REQUEST.QueryString("INTGLASS")
EXTGLASSDOOR = REQUEST.QueryString("EXTGLASSDOOR")
INTGLASSDOOR = REQUEST.QueryString("INTGLASSDOOR")
FSTYLE = REQUEST.QueryString("FSTYLE")
SILLTYPE = REQUEST.QueryString("SILLTYPE")
SPACERCOLOUR = REQUEST.QueryString("SPACERCOLOUR")
AWNSTYLE = REQUEST.QueryString("AWNSTYLE")
LOUVERSTYLE = REQUEST.QueryString("LOUVERSTYLE")


PANELPUNCH = REQUEST.QueryString("PANELPUNCH")
R3VentSize = REQUEST.QueryString("R3VENTSIZE")
if R3VentSize = "" then
	R3VentSize = 0
end if

MaxHoist = REQUEST.QueryString("MaxHoist")
If MaxHoist = "" then
	MaxHoist = 0
End If
VStockLength = REQUEST.QueryString("VStockLength")
If VStockLength = "" then
	VStockLength = 0
End If
JobStatus = REQUEST.QueryString("JobStatus")


RLIST = REQUEST.QueryString("RLIST")
IMPREC = REQUEST.QueryString("IMPREC")
IMPADD = REQUEST.QueryString("IMPADD")
IMPTAX = REQUEST.QueryString("IMPTAX")
EXPTAX = REQUEST.QueryString("EXPTAX")


If FLOORS = "" then
	FLOORS = 0
End If
EXT_COLOUR = REQUEST.QueryString("EXT_COLOUR")
INT_COLOUR = REQUEST.QueryString("INT_COLOUR")
COMPLETED = REQUEST.QueryString("COMPLETED")


ONSITEDATE=REQUEST.QueryString("ONSITEDATE")
STOPCOLOR=REQUEST.QueryString("STOPCOLOR")
NODOORS=REQUEST.QueryString("NODOORS")
NOAWNINGS=REQUEST.QueryString("NOAWNINGS")
SCREEN=REQUEST.QueryString("SCREEN")
SDSILL = REQUEST.QueryString("SDsill")
SWSILL = REQUEST.QueryString("SWsill")

SThickSpacer=REQUEST.QueryString("SThickSpacer")
WThickSpacer=REQUEST.QueryString("WThickSpacer")
OVThickSpacer=REQUEST.QueryString("OVThickSpacer")
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="AllJobsENTER.asp" target="_self">Enter Job</a>

    </div>


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


	if FSTYLE = "" then
		FSTYLE = 0
	end if
	
	if WThickSpacer = "" then
		WTHICK = 0
		SUSPACER = 0
	else
	    wfields  = SPLIT(WThickSpacer,"|")
		WTHICK = wfields(0)
		SUSPACER = wfields(1)
	end if
	if SThickSpacer = "" then
		STHICK = 0
		SWSPACER=0
	else
	    sfields  = SPLIT(SThickSpacer,"|")
		STHICK = sfields(0)
		SWSPACER=sfields(1)
	end if
	if OVThickSpacer = "" then
		ATHICK = 0
		OVSPACER = 0
	else
	    afields  = SPLIT(OVThickSpacer,"|")
		ATHICK = afields(0)
		OVSPACER = afields(1)
	end if

	
	If COMPLETED = "on" then
		COMPLETED = TRUE
	Else
		COMPLETED = FALSE
	End If
	If NODOORS = "on" then
		NODOORS = TRUE
	Else
		NODOORS = FALSE
	End If
	If NOAWNINGS = "on" then
		NOAWNINGS = TRUE
	Else
		NOAWNINGS = FALSE
	End If


	Set rs = Server.CreateObject("adodb.recordset")
	'strSQL = "INSERT INTO Z_Jobs ([JOB], [JOB_NAME], [JOB_ADDRESS], [JOB_CITY], [MATERIAL], [FLOORS], [EXT_COLOUR], [INT_COLOUR], [SILL], [COMPLETED], [MANAGER]) VALUES( '" & JOB & "', '" & JOB_NAME &  "', '" & JOB_ADDRESS & "', '" & JOB_CITY & "', '" & MATERIAL & "', '" & FLOORS & "', '" & EXT_COLOUR & "', '" & INT_COLOUR & "', '" & SILL & "', " & COMPLETED & ", '" & MANAGER & "')"
	'strSQL = "SELECT * FROM Z_Jobs WHERE ID=-1"
	strSQL = "SELECT * FROM Z_Jobs WHERE JOB = '" & JOB & "'"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection
	
	if (rs.eof and Len(JOB)=3) then

		rs.AddNew
		rs.Fields("JOB") = JOB
		rs.Fields("PARENT") = PARENT
		rs.Fields("JOB_NAME") = JOB_NAME
		rs.Fields("JOB_ADDRESS") = JOB_ADDRESS
		rs.Fields("JOB_COUNTRY") = JOB_COUNTRY
		rs.Fields("JOB_CITY") = JOB_CITY
		rs.Fields("MATERIAL") = MATERIAL
		rs.Fields("FLOORS") = FLOORS
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
		rs.Fields("MaxHoist") = MaxHoist
		rs.Fields("VStockLength") = VStockLength
		rs.Fields("JobStatus") = JobStatus
		rs.Fields("StopColor") = STOPCOLOR
		rs.Fields("NoDoors") = NODOORS
		rs.Fields("NoAwnings") = NOAWNINGS
		rs.Fields("Screen") = SCREEN
		rs.Fields("CasAwnThick") = ATHICK
		rs.Fields("AwnStyle") = AWNSTYLE
		rs.Fields("ExtGLass") = EXTGLASS
		rs.Fields("ExtGlassDoor") = EXTGLASSDOOR
		rs.Fields("FRAMESTYLE") = FSTYLE
		rs.Fields("IntGlass") = INTGLASS
		rs.Fields("IntGlassDoor") = INTGLASSDOOR
		rs.Fields("Louverstyle") = LOUVERSTYLE
		rs.Fields("OVSPACER") = OVSPACER
		rs.Fields("SillType") = SILLTYPE		
		rs.Fields("SpacerColour") = SPACERCOLOUR
		rs.Fields("SwingDoorThick") = STHICK
		rs.Fields("SUSPACER") = SUSPACER
		rs.Fields("FixWindowThick") = WTHICK
		rs.Fields("SDSILL") = SDSILL
		rs.Fields("SWSILL") = SWSILL
		rs.Fields("SWSPACER") = SWSPACER
		
		if isDate(ONSITEDATE) then
			rs.Fields("OnSiteDate") = ONSITEDATE
		end if		

		If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
		rs.update

		Call StoreID1(isSQLServer, rs.Fields("ID"))
	
	else
		if Len(Job) = 3 then
			Exists = True
			ErrorComment = "Already Exists in the Database, Please Edit the existing record"
		else
			Exists = True
			ErrorComment = "Please enter a 3 Letter Job Code"
		End if													   									
	End if
	DbCloseAll
End If
End Function

%>

<% 
if Password = Passkey then
%> 
<form id="conf" title="Added" class="panel" name="conf" action="index.html#_Job" method="GET" target="_self" selected="true" >              
	
        <h2>Parent Job Addition</h2> 
<ul id="Report" title="Parent Job Addition" selected="true">
	<%
	if Exists = True then
	%>
		<li><% response.write "JOB: " & JOB %></li>
		<li><% response.write ErrorComment %></li>
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
		<li><% response.write "Job Status: " & JobStatus %></li>
		<li><% response.write "COMPLETED: " & COMPLETED %></li>
	<%
	end if
	%>	
	</ul>

        <BR>
         <a class="whiteButton" href="AllJobsReport.asp" target="_self"> Back to All Jobs List</a>
            </form>			
<%
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="AllJobsConf.asp?JOB=<%response.write JOB%>&PARENT=<%response.write PARENT%>&JOB_NAME=<%response.write JOB_NAME%>&JOB_ADDRESS=<%response.write JOB_ADDRESS%>&JOB_CITY=<%response.write JOB_CITY%>&JOB_COUNTRY=<%response.write JOB_COUNTRY%>&MATERIAL=<%response.write MATERIAL%>&FLOORS=<%response.write FLOORS%>&COLORMATCH=<%response.write COLORMATCH%>&RLIST=<%response.write RLIST%>&IMPREC=<%response.write IMPREC%>&IMPADD=<%response.write IMPADD%>&IMPTAX=<%response.write IMPTAX%>&EXPTAX=<%response.write EXPTAX%>&MANAGER=<%response.write MANAGER%>&MANAGEREMAIL=<%response.write MANAGEREMAIL%>&ENGINEER=<%response.write ENGINEER%>&ENGINEEREMAIL=<%response.write ENGINEEREMAIL%>&EXT_COLOUR=<%response.write EXT_COLOUR%>&INT_COLOUR=<%response.write INT_COLOUR%>&COMPLETED=<%response.write COMPLETED%>&EXTGLASS=<%response.write EXTGLASS%>&INTGLASS=<%response.write INTGLASS%>&EXTGLASSDOOR=<%response.write EXTGLASSDOOR%>&INTGLASSDOOR=<%response.write INTGLASSDOOR%>&FSTYLE=<%response.write FSTYLE%>&WTHICK=<%response.write WTHICK%>&STHICK=<%response.write STHICK%>&ATHICK=<%response.write ATHICK%>&SUSPACER=<%response.write SUSPACER%>&OVSPACER=<%response.write OVSPACER%>&SILLTYPE=<%response.write SILLTYPE%>&SPACERCOLOUR=<%response.write SPACERCOLOUR%>&AWNSTYLE=<%response.write AWNSTYLE%>&LOUVERSTYLE=<%response.write LOUVERSTYLE%>&GLASSFLUSH=<%response.write GLASSFLUSH%>&BEAUTYSTYLE=<%response.write BEAUTYSTYLE%>&PANELPUNCH=<%response.write PANELPUNCH%>&R3VentSize=<%response.write R3VentSize%>&MaxHoist=<%response.write MaxHoist%>&VStockLength=<%response.write VStockLength%>&JobStatus=<%response.write JobStatus%>&ONSITEDATE=<%response.write ONSITEDATE%>&NODOORS=<%response.write NODOORS%>&NOAWNINGS=<%response.write NOAWNINGS%>&STOPCOLOR=<%response.write STOPCOLOR%>&SCREEN=<%response.write SCREEN%>&SDSILL=<%response.write SDSILL%>&SWSILL=<%response.write SWSILL%>&SThickSpacer=<%response.write SThickSpacer%>&WThickSpacer=<%response.write WThickSpacer%>&OVThickSpacer=<%response.write OVThickSpacer%>" method="post" target="_self" selected="True">
	<ul id="Report" title="Password" selected="true">
<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	
         <a class="whiteButton" href="AllJobsReport.asp#_Profiles" target="_self"> Back</a>
            </form>		
<%
end if
%>


<%

'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>

