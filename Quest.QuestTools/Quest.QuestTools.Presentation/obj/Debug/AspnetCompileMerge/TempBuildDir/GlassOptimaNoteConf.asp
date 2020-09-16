<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Add Note Confirmation page - from GlassOptimaNote.asp Requested by Sasha-->
		 <!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Edited </title>
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

ticket = request.querystring("ticket")
	Select Case ticket
		case "select"
			   Sender = "GlassExportSelect.asp"
		case "active"
			   Sender = "GlassReportActive.asp" 
		case "active"
			   Sender = "GlassReportCompleted.asp" 
		case "multiple"
			   Sender = "GlassOptimaNoteSelect.asp" 
	End Select
%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassOptimaNoteSelect.asp" target="_self">Add Note</a>
    </div>

<form id="conf" title="Glass Edited" class="panel" name="conf" action= " <%response.write Sender%>" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%

GIDList = ""

NOTES = REQUEST.QueryString("NOTES")
PO = REQUEST.QueryString("PO")
EXTWORK = REQUEST.QueryString("EXTORDERNUM")
INTWORK = REQUEST.QueryString("INTORDERNUM")
QTFILE = REQUEST.QueryString("QTFILE")

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

	For Each item in Request.QueryString("GID")
		GID = item

		If Instr(1, GIDList, GID) < 1 Then GIDList = GIDList & GID & ", " 

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * from Z_GLASSDB WHERE ID = " & GID
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

		if NOTES = "" or isNUll(NOTES) then
		else
		rs("Notes") = NOTES
		end if

		if PO = "" or isNUll(PO) then
		else
		rs("PO") = PO
		end if

		if ExtWork = "" or isNUll(ExtWork) then
		else
		rs("ExtOrderNum") = ExtWork
		end if

		if IntWork = "" or isNUll(IntWork) then
		else
		rs("IntOrderNum") = INTWORK
		end if

		if QTFILE = "" or isNUll(QTFILE) then
		else
		rs("QTFILE") = QTFILE
		end if

		rs.update
		rs.close
		set rs = nothing

	Next

	DbCloseAll

End Function

%>

<ul id="Report" title="Added" selected="true">

<%

		Response.Write "<li>Details ADDED:</li>"
		Response.Write "<li> Details Added to : " & GIDList & "</li>"
		Response.Write "<li> Notes: " & NOTES & "</li>"
		Response.Write "<li> Window PO: " & PO & "</li>"
		Response.Write "<li> External Work Order: " & EXTWORK & "</li>"
		Response.Write "<li> Internal Work Order: " & INTWORK & "</li>"
		Response.Write "<li> QTFILE: " & QTFILE & "</li>"

%>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Optima Select</a>
            
            </form>

</body>
</html>

<%

'DBConnection.close
'set DBConnection=nothing
%>

