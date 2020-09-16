<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created Janaury 28th, 2015 - by Michael Bernholtz --> 
<!--Add QT Confirmation page - from GlassOptimaQT.asp Requested by Sasha-->
<!--Based on Comma and entered ids instead of selection -->
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

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassOptimaNoteSelect2.asp" target="_self">Add Key </a>
    </div>
    
      
    
<form id="conf" title="Glass Edited" class="panel" name="conf" action= "GlassOptimaNoteSelect2.asp" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%
GIDList = ""
QT = REQUEST.QueryString("QT")
Notes = REQUEST.QueryString("Notes")
PO = REQUEST.QueryString("PONum")
ExtWork = REQUEST.QueryString("ExtOrderNum")
IntWork = REQUEST.QueryString("IntOrderNum")
QTFile = REQUEST.QueryString("QTFile")
IDLIST = Request.QueryString("IDList")
IDLIST = IDLIST & ","
IDLIST = Replace(IDlist," ","")
GIDLIST = IDLIST
counter =1

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
			IDLIST = GIDLIST
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	Do While INSTR(1, IDLIST, ",",1)
		counter =counter+1
		CommaPlace = Instr(1, IDLIST, ",",1)
		GID = LEFT(IDLIST,CommaPlace-1) 

		IDLIST = Right(IDLIST, LEN(IDLIST)- CommaPlace)

		'Set Glass Inventory Update Statement

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * from Z_GLASSDB WHERE ID = " & GID
		'DebugMsg(strSQL & "<br/>")
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection

		If NOTES = "" or isNUll(NOTES) Then
			'DebugMsg("No: " & GID & " SQL: " & isSQLServer & "<br/>")
		Else
			'DebugMsg("No: " & GID & " Note: " & NOTES & " SQL: " & isSQLServer & "<br/>")
			rs("Notes") = NOTES
		End If
	
		If PO = "" or isNUll(PO) Then
		Else
			rs("PO") = PO
		End If

		If ExtWork = "" or isNUll(ExtWork) Then
		Else
			rs("ExtOrderNum") = ExtWork
		End If

		If IntWork = "" or isNUll(IntWork) Then
		Else
			rs("IntOrderNum") = INTWORK
		End If

		If QTFILE = "" or isNUll(QTFILE) Then
		Else
			rs("QTFILE") = QTFILE
		End If

		rs.update
		rs.close
		set rs = nothing

	loop

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

