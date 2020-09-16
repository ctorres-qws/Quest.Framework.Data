<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
<!--AllJobs Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->
<!-- Updated for Global Variables Updates by Annabel Ramirez August 2019 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Summary Edited </title>
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

JID = request.querystring("JID")

Passkey = "QUEJAY"
Password = UCASE(TRIM(Request.Form("pwd")))

JOB = UCASE(REQUEST.QueryString("JOB"))
PARENT = UCASE(REQUEST.QueryString("PARENT"))

JOB_NAME = REQUEST.QueryString("JOB_NAME")

JOB_ADDRESS = REQUEST.QueryString("JOB_ADDRESS")
JOB_CITY = REQUEST.QueryString("JOB_CITY")
JOB_COUNTRY = REQUEST.QueryString("JOB_COUNTRY")
MATERIAL = REQUEST.QueryString("MATERIAL")
'SILL = REQUEST.QueryString("SILL")
FLOORS = REQUEST.QueryString("FLOORS")
COLORMATCH = REQUEST.QueryString("COLORMATCH")

RLIST = REQUEST.QueryString("RLIST")
IMPREC = REQUEST.QueryString("IMPREC")
IMPADD = REQUEST.QueryString("IMPADD")
IMPTAX = REQUEST.QueryString("IMPTAX")
EXPTAX = REQUEST.QueryString("EXPTAX")

MANAGER = REQUEST.QueryString("MANAGER")
MANAGEREMAIL = REQUEST.QueryString("MANAGEREMAIL")
ENGINEER = REQUEST.QueryString("ENGINEER")
ENGINEEREMAIL = REQUEST.QueryString("ENGINEEREMAIL")
EXT_COLOUR = REQUEST.QueryString("EXT_COLOUR")
INT_COLOUR = REQUEST.QueryString("INT_COLOUR")
COMPLETED = REQUEST.QueryString("COMPLETED")
EXTGLASS = REQUEST.QueryString("EXTGLASS")
INTGLASS = REQUEST.QueryString("INTGLASS")
EXTGLASSDOOR = REQUEST.QueryString("EXTGLASSDOOR")
INTGLASSDOOR = REQUEST.QueryString("INTGLASSDOOR")
FSTYLE = REQUEST.QueryString("FSTYLE")
WTHICK = REQUEST.QueryString("WTHICK")
STHICK = REQUEST.QueryString("STHICK")
ATHICK = REQUEST.QueryString("ATHICK")
SUSPACER = REQUEST.QueryString("SUSPACER")
OVSPACER = REQUEST.QueryString("OVSPACER")
SILLTYPE = REQUEST.QueryString("SILLTYPE")
SPACERCOLOUR = REQUEST.QueryString("SPACERCOLOUR")
AWNSTYLE = REQUEST.QueryString("AWNSTYLE")
LOUVERSTYLE = REQUEST.QueryString("LOUVERSTYLE")
GLASSFLUSH = REQUEST.QueryString("GLASSFLUSH")
BEAUTYSTYLE = REQUEST.QueryString("BEAUTYSTYLE")
PANELPUNCH = REQUEST.QueryString("PANELPUNCH")
R3VentSize = REQUEST.QueryString("R3VENTSIZE")
if R3VentSize = "" then
	R3VentSize = 0
end if

MaxHoist = REQUEST.QueryString("MaxHoist")
if MaxHoist = "" then
	MaxHoist = 0
end if
VStockLength = REQUEST.QueryString("VStockLength")
if VStockLength = "" then
	VStockLength = 0
end if
JobStatus = REQUEST.QueryString("JobStatus")


if FSTYLE = "" then
	FSTYLE = 0
end if


if FLOORS = "" then
	FLOORS = 0
end if

currentDate = Date()
ONSITEDATE=REQUEST.QueryString("OnSiteDate")
NODOORS = REQUEST.QueryString("NODOORS")
NOAWNINGS = REQUEST.QueryString("NOAWNINGS")
STOPCOLOR = REQUEST.QueryString("STOPCOLOR")
SCREEN = REQUEST.QueryString("SCREEN")
SDSILL = REQUEST.QueryString("SDSILL")
SWSILL = REQUEST.QueryString("SWSILL")

SThickSpacer=REQUEST.QueryString("SThickSpacer")
WThickSpacer=REQUEST.QueryString("WThickSpacer")
OVThickSpacer=REQUEST.QueryString("OVThickSpacer")

'ADDITION OF SHIPPING LABEL COLOR FOR SCAN TO PRINT JUN, 2020 BY CTORRES
ShippingLabelColor = REQUEST.QueryString("ShippingLabelColor")

%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="AllJobsEditForm.asp?JID=<% response.write Jid %>" target="_self">Edit Job</a>
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

			'Set Glass Inventory Update Statement
			'StrSQL = FixSQLCheck("UPDATE Z_JOBS  SET [JOB]='"& JOB & "', [JOB_NAME]='" & JOB_NAME & "', [JOB_ADDRESS]='" & JOB_ADDRESS & "', JOB_CITY= '" & JOB_CITY & "', [MATERIAL]='" & MATERIAL & "', [FLOORS]='" & FLOORS & "', [EXT_COLOUR]= '" & EXT_COLOUR & "', [INT_COLOUR]= '" & INT_COLOUR & "', [SILL]= '" & SILL & "', [MANAGER]= '" & MANAGER & "', [COMPLETED]= " & COMPLETED & " WHERE ID = " & JID, isSQLServer)
			'Get a Record Set
			'Set RS = DBConnection.Execute(strSQL)
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

		
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_Jobs WHERE ID = " & JID
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

		rs.Fields("JOB") = JOB
		rs.Fields("PARENT") = PARENT
		rs.Fields("JOB_NAME") = JOB_NAME
		rs.Fields("JOB_ADDRESS") = JOB_ADDRESS
		rs.Fields("JOB_CITY") = JOB_CITY
		rs.Fields("JOB_COUNTRY") = JOB_COUNTRY
		rs.Fields("FLOORS") = INT(FLOORS)
		rs.Fields("COLORMATCH") = COLORMATCH
		
		rs.Fields("RecipientList") = RLIST
		rs.Fields("ImporterRecord") = IMPREC
		rs.Fields("ImporterAddress") = IMPADD
		rs.Fields("ImporterTaxID") = IMPTAX
		rs.Fields("ExporterTaxID") = EXPTAX
		
		rs.Fields("MATERIAL") = MATERIAL
		rs.Fields("FRAMESTYLE") = FSTYLE
		'rs.Fields("SILL") = SILL
		rs.Fields("EXT_Colour") = EXT_COLOUR	
		rs.Fields("INT_Colour") = INT_COLOUR
		rs.Fields("Completed") = COMPLETED
		rs.Fields("ENGINEER") = ENGINEER
		rs.Fields("ENGINEEREMAIL") = ENGINEEREMAIL
		rs.Fields("Manager") = MANAGER
		rs.Fields("ManagerEmail") = MANAGEREMAIL
		rs.Fields("ExtGLass") = EXTGLASS
		rs.Fields("IntGlass") = INTGLASS
		rs.Fields("ExtGlassDoor") = EXTGLASSDOOR
		rs.Fields("IntGlassDoor") = INTGLASSDOOR
		rs.Fields("FixWindowThick") = CINT(WTHICK)
		rs.Fields("SwingDoorThick") = CINT(STHICK)
		rs.Fields("CasAwnThick") = CINT(ATHICK)
		rs.Fields("SUSPACER") = CINT(SUSPACER)
		rs.Fields("OVSPACER") = CINT(OVSPACER)
		rs.Fields("SWSPACER") = CINT(SWSPACER)
		rs.Fields("SillType") = SILLTYPE
		rs.Fields("SpacerColour") = SPACERCOLOUR
		rs.Fields("AwnStyle") = AWNSTYLE
		rs.Fields("Louverstyle") = LOUVERSTYLE
		rs.Fields("GlassFlush") = GLASSFLUSH
		rs.Fields("BeautyStyle") = BEAUTYSTYLE
		rs.Fields("PanelPunch") = PANELPUNCH
		rs.Fields("R3VentSize") = INT(R3VENTSIZE)
		
		rs.Fields("MaxHoist") = MaxHoist
		rs.Fields("VStockLength") = VStockLength
		rs.Fields("JobStatus") = JobStatus
		
		rs.Fields("ModifiedDate") = currentDate
		
		if isDate(ONSITEDATE) then
			rs.Fields("OnSiteDate") = ONSITEDATE
		end if	
		
		rs.Fields("StopColor") = STOPCOLOR
		rs.Fields("NoDoors") = NODOORS
		rs.Fields("NoAwnings") = NOAWNINGS
		rs.Fields("Screen") = SCREEN	
		rs.Fields("SDSILL") = SDSILL
		rs.Fields("SWSILL") = SWSILL
		rs.Fields("ShippingLabelColor") = "#" & ShippingLabelColor
		rs.UPDATE

		DbCloseAll
	End If
End Function

%>

<% 
if Password = Passkey then
%> 
<form id="conf" title="Job Summary Edited" class="panel" name="conf" action="index.html#_Job" method="GET" target="_self" selected="true" >              
	
        <h2>Job Summary Edited</h2> 
	<ul id="Report" title="Job Summary Edited" selected="true">

<%	
		Response.Write "<li> JOB: " & JOB & "</li>"
		Response.Write "<li> PARENT: " & PARENT & "</li>"
		Response.Write "<li> FULL NAME: " & JOB_NAME & "</li>"
		Response.Write "<li> ADDRESS: " & JOB_ADDRESS & "</li>"
		Response.Write "<li> CITY: " & JOB_CITY & "</li>"
		Response.Write "<li> COUNTRY: " & JOB_COUNTRY & "</li>"
		Response.Write "<li> MATERIAL: " & MATERIAL & "</li>"
		Response.Write "<li> MANAGER: " & MANAGER & "</li>"
		Response.Write "<li> FLOORS: " & FLOORS & "</li>"
		Response.Write "<li> Recipient List: " & RLIST & "</li>"
		Response.Write "<li> EXTERIOR COLOUR: " & EXT_COLOUR & "</li>"
		Response.Write "<li> INTERIOR COLOUR: " & INT_COLOUR & "</li>"
		Response.Write "<li> EXTERIOR GLASS DOOR: " & EXTGLASSDOOR & "</li>"
		Response.Write "<li> INTERIOR  GLASS DOOR: " & INTGLASSDOOR & "</li>"
		Response.Write "<li> HBAR Color Match: " & COLORMATCH & "</li>"		
		Response.Write "<li> Job Status: " & JobStatus & "</li>"		
		Response.Write "<li> COMPLETED: " & COMPLETED & "</li>"	
		Response.Write "<li> Shipping Label Color: " & ShippingLabelColor & "</li>"	
		
%>	

	</ul>
        <BR>
         <a class="whiteButton" href="AllJobsReport.asp#_Profiles" target="_self"> Back</a>
            </form>			
<%
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="AllJobsEditConf.asp?JID=<%response.write JID%>&JOB=<%response.write JOB%>&PARENT=<%response.write PARENT%>&JOB_NAME=<%response.write JOB_NAME%>&JOB_ADDRESS=<%response.write JOB_ADDRESS%>&JOB_CITY=<%response.write JOB_CITY%>&JOB_COUNTRY=<%response.write JOB_COUNTRY%>&MATERIAL=<%response.write MATERIAL%>&FLOORS=<%response.write FLOORS%>&ShippingLabelColor=<%response.write ShippingLabelColor%>&COLORMATCH=<%response.write COLORMATCH%>&RLIST=<%response.write RLIST%>&IMPREC=<%response.write IMPREC%>&IMPADD=<%response.write IMPADD%>&IMPTAX=<%response.write IMPTAX%>&EXPTAX=<%response.write EXPTAX%>&MANAGER=<%response.write MANAGER%>&MANAGEREMAIL=<%response.write MANAGEREMAIL%>&ENGINEER=<%response.write ENGINEER%>&ENGINEEREMAIL=<%response.write ENGINEEREMAIL%>&EXT_COLOUR=<%response.write EXT_COLOUR%>&INT_COLOUR=<%response.write INT_COLOUR%>&COMPLETED=<%response.write COMPLETED%>&EXTGLASS=<%response.write EXTGLASS%>&INTGLASS=<%response.write INTGLASS%>&EXTGLASSDOOR=<%response.write EXTGLASSDOOR%>&INTGLASSDOOR=<%response.write INTGLASSDOOR%>&FSTYLE=<%response.write FSTYLE%>&WTHICK=<%response.write WTHICK%>&STHICK=<%response.write STHICK%>&ATHICK=<%response.write ATHICK%>&SUSPACER=<%response.write SUSPACER%>&OVSPACER=<%response.write OVSPACER%>&SILLTYPE=<%response.write SILLTYPE%>&SPACERCOLOUR=<%response.write SPACERCOLOUR%>&AWNSTYLE=<%response.write AWNSTYLE%>&LOUVERSTYLE=<%response.write LOUVERSTYLE%>&GLASSFLUSH=<%response.write GLASSFLUSH%>&BEAUTYSTYLE=<%response.write BEAUTYSTYLE%>&PANELPUNCH=<%response.write PANELPUNCH%>&R3VentSize=<%response.write R3VentSize%>&MaxHoist=<%response.write MaxHoist%>&VStockLength=<%response.write VStockLength%>&JobStatus=<%response.write JobStatus%>&ONSITEDATE=<%response.write ONSITEDATE%>&NODOORS=<%response.write NODOORS%>&NOAWNINGS=<%response.write NOAWNINGS%>&STOPCOLOR=<%response.write STOPCOLOR%>&SCREEN=<%response.write SCREEN%>&SDSILL=<%response.write SDSILL%>&SWSILL=<%response.write SWSILL%>&SThickSpacer=<%response.write SThickSpacer%>&WThickSpacer=<%response.write WThickSpacer%>&OVThickSpacer=<%response.write OVThickSpacer%>" method="post" target="_self" selected="true">

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

            
    
</body>
</html>


