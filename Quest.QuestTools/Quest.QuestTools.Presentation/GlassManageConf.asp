<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Edit Confirmation Page for Glass Items-->
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

Gid = request.querystring("Gid")
ticket = Request.QueryString("ticket")

Select Case ticket
	case "active"
		returnEmail = "GlassToolViewActive.asp"
	case "optima"
		returnEmail = "GlassToolViewOptima.asp"
	case "waiting"
		returnEmail = "GlassToolViewWait.asp"
	case "received"
		returnEmail = "GlassToolViewReceived.asp"
	case "completed"
		returnEmail = "GlassToolViewCompleted.asp"
	case "shipped"
		returnEmail = "GlassToolViewShipped.asp"
	case Else
		returnEmail = "GlassManage.asp"
End select
%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassManageForm.asp?GID=<% response.write Gid %>" target="_self">Edit Glass</a>
    </div>

<form id="conf" title="Glass Edited" class="panel" name="conf" action="<%response.write returnEmail%>" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%

BARCODE = "GT" & ID

FLOOR = REQUEST.QueryString("FLOOR")
TAG = REQUEST.QueryString("TAG")
JOB = REQUEST.QueryString("PROJECT")
CUSTOMER = JOB
DEPARTMENT = REQUEST.QueryString("DEPARTMENT")
WIDTH = REQUEST.QueryString("WIDTH")
HEIGHT = REQUEST.QueryString("HEIGHT")

ONEMAT = REQUEST.QueryString("ONEMAT")
TWOMAT = REQUEST.QueryString("TWOMAT")
ONESPAC = REQUEST.QueryString("ONESPAC")


ORDERBY = REQUEST.QueryString("orderBy")
PoNum = REQUEST.QueryString("PoNum")
QTFile = REQUEST.QueryString("QTFile")
NOTES = REQUEST.QueryString("NOTES")
AIR = REQUEST.QueryString("AIR")	
  
'Added for Jody and Ruslan - Glass Tools system	   
ExtorderNum = REQUEST.QueryString("ExtOrderNum")	
IntorderNum = REQUEST.QueryString("IntorderNum")	
BackorderFlag = REQUEST.QueryString("BackorderFlag")	

ExtFrom = REQUEST.QueryString("ExtFrom")	
IntFrom= REQUEST.QueryString("IntFrom")

' Added April 2015 for IVAN and SASHA
ExtMethod = REQUEST.QueryString("EXTMethod")
If EXTMethod = "ALREADY-HAVE" Then
	ONEMAT = "X" & ONEMAT & "X"
End If

IntMethod = REQUEST.QueryString("INTMethod")
If INTMethod = "ALREADY-HAVE" Then
	TWOMAT = "X" & TWOMAT & "X"
End If

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

	'Set Glass Inventory Update Statement
	StrSQL = "UPDATE Z_GLASSDB  SET [JOB]='"& JOB & "', [FLOOR]='" & FLOOR & "', [TAG]='" & TAG & "', Customer= '" & CUSTOMER & "', [DEPARTMENT]='" & DEPARTMENT & "', [DIM X]='" & WIDTH & "', [DIM Y]= '" & HEIGHT & "', [1 MAT]= '" & ONEMAT & "', [2 MAT]= '" & TWOMAT & "', [1 SPAC]= '" & ONESPAC & "', [ORDERBY]= '" & ORDERBY & "', [ORDERFor]= '" & ORDERFor & "', [PO]= '" & PoNum & "', [NOTES]= '" & NOTES & "', [AIR]= '" & AIR & "', [QTFile]= '" & QTfile & "', [Extordernum]= '" & ExtorderNum & "', [IntOrderNum]= '" & IntorderNum & "', [BackorderFlag]= '" & BackorderFlag & "', [ExtFrom]= '" & ExtFrom & "', [IntFrom]= '" & IntFrom & "', [ExtMethod]= '" & ExtMethod & "', [IntMethod]= '" & IntMethod & "'  WHERE ID = " & GID
	'Get a Record Set
	Set RS = DBConnection.Execute(strSQL)

	DbCloseAll

End Function

%>

<ul id="Report" title="Added" selected="true">

<%

		Response.Write "<li>Inventory GLASS Edited:</li>"
		Response.Write "<li> Floor: " & FLOOR & "</li>"
		Response.Write "<li> Tag: " & TAG & "</li>"
		Response.Write "<li> Job: " & JOB & "</li>"
		Response.Write "<li> Department: " & DEPARTMENT & "</li>"
		Response.Write "<li> Width: " & WIDTH & "</li>"
		Response.Write "<li> 1 MAT: " & ONEMAT & "</li>"
		Response.Write "<li> SPACER: " & ONESPAC & "</li>"
		Response.Write "<li> 2 MAT: " & TWOMAT & "</li>"
		Response.Write "<li> Height: " & HEIGHT & "</li>"
		Response.Write "<li> Customer: " & CUSTOMER & "</li>"
		Response.Write "<li> Ordered By: " & ORDERBY & "</li>"
		Response.Write "<li> Ordered For: " & ORDERFor & "</li>"
		Response.Write "<li> Notes: " & NOTES & "</li>"
		Response.Write "<li> Air/Argon: " & AIR & "</li>"
		Response.Write "<li> Window PO: " & PoNum & "</li>"
		Response.Write "<li> External Glass Work order: " & ExtorderNum & "</li>"
		Response.Write "<li> Internal Glass Work order: " & IntorderNum & "</li>"
		Response.Write "<li> QT File: " & QTFile & "</li>"

		If not isNull(BackorderFlag) Then
			Response.Write "<li> Back Order: " & BackOrderFlag & "</li>"
		End If

%>
        <BR>
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
		<%
		' Edit next and Previous form added at request of Tomos March 2015
		idnext = "GlassManageForm.asp?GID=" & Gid + 1
		idprevious = "GlassManageForm.asp?GID=" & Gid - 1

		%>

		<a class="greenButton" href ='<% response.write idnext %>'> Edit Next (ID:<%response.write Gid + 1 %>)</a><BR>
		<a class="lightblueButton" href = '<% response.write idprevious %>'> Edit Previous (ID:<%response.write Gid -1 %>) </a><BR>

            </form>

</body>
</html>
<%
'DBConnection.close
'set DBConnection=nothing
%>

