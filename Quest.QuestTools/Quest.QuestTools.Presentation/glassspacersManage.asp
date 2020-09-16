<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 16th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page SPACERCONF.asp -->
		 <!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Spacers</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job / Colour</a>
        </div>

              <form id="enter" title="Manage Glass" class="panel" name="managespacers" action="glassspacersManage.asp" method="GET" target="_self" selected="true">
<%
	SPACER = request.querystring("SPACER")
	OT = request.querystring("OT")
	SPACERID = request.querystring("ID")
	UPDATEFLAG = request.querystring("UPDATE")

	If UPDATEFLAG = 1 Then
		FLAG = 0
	Else 
		Flag =1
	End If

%>
		<h2>Manage SPACER</h2>
<%

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
			If UPDATEFLAG = 1 Then
				SPACER = "0"
				OT = "0"
				SPACERID = "0"
				UPDATEFLAG = "0"
			End If
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	If  SPACER = "0" AND OT = "0" AND SPACERID ="0" Then
		response.write ""
	Else
		If UPDATEFLAG = 1 Then
			Set rs7 = Server.CreateObject("adodb.recordset") 
			strSQL7 = "UPDATE XQSU_OTSpacer SET Spacer='" & SPACER& "', OT='" & OT & "'  WHERE ID = " & SPACERID

			rs7.Cursortype = GetDBCursorType
			rs7.Locktype = GetDBLockType
			rs7.Open strSQL7, DBConnection

			If gi_Mode = c_MODE_ACCESS Or gi_Mode = c_MODE_SQL_SERVER Then
				SPACER = "0"
				OT = "0"
				SPACERID = "0"
				UPDATEFLAG = "0"
			End If

		End if
	End If

DbCloseAll

End Function


		If  SPACER = "0" AND OT = "0" AND SPACERID ="0" Then
			Response.write ""
		Else 
%>
			<fieldset>

				<!--Entry form to edit existing Type and OT and then replace the old one in the database-->

			<div class="row">
				<% response.write"<input type='hidden' name='ID' id='SPACERID' value = '"& SPACERID & "' readonly = 'readonly'> "
				%>
			<div>
			<div class="row">
				<% response.write"<input type='hidden' name='UPDATE' id='UPDATE' value = '"& FLAG & "' readonly = 'readonly'> "
				%>			
			<div>
			<div class="row">
				<label> Type: </label>
				<% response.write "<input type='text' name='SPACER' value = '" & SPACER & "' id='SPACER' > "
				%>
				<ul>
				<li>Note: Spacer must be a number</li>
				</ul>
			</div>
	
			<div class="row">
				<label> OT: </label>
				<%response.write "<input type='text' name='OT' value = '" & OT & "' id='OT' >"
				%>
				<ul>
				<li>Note: Please List OT in format "Overall Inches (Xmm / Xmm / Xmm) </li>
				</ul>
			</div>
	
			<a class="whiteButton" href="javascript:managespacers.submit()">Submit</a>
		
			
			<!-- SUBMIT BUTTON has to change to Edit the existing field rather than create new-->
</fieldset>
<%
		End If
%>

<ul id="Profiles" title="Glass" selected="true">

<%

response.write "<li class = 'group'> Current Spacers:</li>" 
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp
DBConnection.Open DSN

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "Select * FROM XQSU_OTSpacer ORDER BY SPACER ASC"
rs6.Cursortype = GetDBCursorType
rs6.Locktype = GetDBLockType
rs6.Open strSQL6, DBConnection

rs6.movefirst
do while not rs6.eof
	response.write "<li><a href='glassspacersManage.asp?SPACER=" & rs6("Spacer") & "&OT=" & Server.UrlEncode(rs6("OT") & "") & "&ID=" & rs6("ID") & "&UDPATE=0' target='_self'> Spacer: " & rs6("Spacer") & " - " & rs6("OT") & "</a></li>"
	rs6.movenext
loop
	
%>
	</ul>
	</form>

<%
rs6.close
set rs6=nothing

DBConnection.close
set DBConnection=nothing
%>
</body>
</html>
