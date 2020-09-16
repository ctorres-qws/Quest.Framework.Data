<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page GLASSSpandrelCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Spandrel Color</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>

              <form id="enter" title="Manage Spandrel Color" class="panel" name="managetypes" action="glassSpandrelManage.asp" method="GET" target="_self" selected="true">
<%
Dim rs5
		Code = request.querystring("Code")
		Description = request.querystring("Description")
		Job = request.querystring("Job")
		Notes = request.querystring("Notes")
		Active = request.querystring("Active")
		if Active = "True" or Active = "on"  then
			Active = TRUE
		else
			Active = FALSE
		end if

		SCID = request.querystring("SCID") 
		UPDATEFLAG = request.querystring("UPDATE")
		IF UPDATEFLAG = 1 then
			Flag =0
		else 
			Flag =1
		end if

%>
		<h2>Manage Glass</h2>

<%

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)

			If UPDATEFLAG = 1 Then
				CODE= "0"
				DESCRIPTION = "0"
				SCID = "0"
				UPDATEFLAG = "0"
			End If
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	If  CODE = "0" AND DESCRIPTION = "0" AND SCID ="0" Then
		response.write ""
	Else 
		If UPDATEFLAG = 1 Then
			strSQL6 = FixSQLCheck("UPDATE Y_COLOR_SPANDREL SET Code='" & Code & "', Description='" & Description & "', Job='" & Job & "', Notes='" & Notes & "', Active= " & Active & "  WHERE ID = " & SCID, isSQLServer)
			DBConnection.Execute strSQL6
			If gi_Mode = c_MODE_ACCESS Or gi_Mode = c_MODE_SQL_SERVER Then
				CODE= "0"
				DESCRIPTION = "0"
				SCID = "0"
				UPDATEFLAG = "0"
			End If
		End If
	End If

'DbCloseAll

End Function

		if  CODE = "0" AND DESCRIPTION = "0" AND SCID ="0" then
		response.write ""
		else 
%>

			<fieldset>

				<!--Entry form to edit existing Type and Description and then replace the old one in the database-->

				<input type='hidden' name='SCID' id='SCID' value = '<%response.write SCID %>'> 
				<input type='hidden' name='UPDATE' id='UPDATE' value = '<%response.write FLAG %>'> 

			<div class="row">
				<label> Code: </label>
				<input type='text' name='Code' value = '<% response.write Code %>' id='Code' > 

			</div>

			<div class="row">
				<label> Description: </label>
				<input type='text' name='Description' value = '<% response.write Description %>' id='Description' >
			</div>
						<div class="row">
				<label> Job: </label>
				<input type='text' name='Job' value = '<% response.write Job %>' id='Job' >
			</div>

			<div class="row">
				<label> Notes: </label>
				<input type='text' name='Notes' value = '<% response.write Notes %>' id='Notes' >
			</div>
						<div class="row">
				<label> Active: </label>
				<input type="checkbox" name='Active' id='Active'  <% if ACTIVE = True then response.write "checked" End if %>>
			</div>

			<a class="whiteButton" href="javascript:managetypes.submit()">Submit</a>

			<!-- SUBMIT BUTTON has to change to Edit the existing field rather than create new-->
</fieldset>
		<% end if
		%>

<ul id="Profiles" title="Glass" selected="true">

<%

response.write "<li class = 'group'> Current Spandrel Colours:</li>" 

Set rs5 = Server.CreateObject("adodb.recordset")
strSQ5L = "Select * FROM Y_COLOR_SPANDREL ORDER BY Code ASC"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQ5L, DBConnection

rs5.movefirst
do while not rs5.eof
	response.write "<li><a href='glassSpandrelManage.asp?Code=" & rs5("Code") & "&Description=" & rs5("Description")& "&Job=" & rs5("Job")&" &Notes=" & rs5("Notes")& "&Active=" & rs5("Active") & "&SCID=" & rs5("ID") & "&UDPATE=0' target='_self'> ID: " & rs5("CODE") & " - " & rs5("Description") & " - Active: " & rs5("Active") & "</a></li>"
	
rs5.movenext
loop

%>
	</ul>
	</form>

<%
rs5.close
set rs5=nothing

DBConnection.close
set DBConnection=nothing
%>
</body>
</html>
