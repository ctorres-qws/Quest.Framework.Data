<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page GLASStypeCONF.asp -->
		 <!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Glass Types</title>
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

              <form id="enter" title="Manage Glass" class="panel" name="managetypes" action="glasstypesManage.asp" method="GET" target="_self" selected="true">
<%
	GLASSTYPE = request.querystring("GlassType")
	DESCRIPTION = request.querystring("Description")
	SHOPCODE = request.querystring("ShopCode")
	STATUS = request.querystring("Status")
	GLASSTYPEID = request.querystring("ID")
	UPDATEFLAG = request.querystring("UPDATE")

	If UPDATEFLAG = 1 Then
		FLAG = 0
	Else
		Flag =1
	End If

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
			GLASSTYPE = "0"
			DESCRIPTION = "0"
			SHOPCODE = "0"
			STATUS = "0"
			GLASSTYPEID = "0"
			UPDATEFLAG = "0"
		End If
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	If  GLASSTYPE = "0" AND DESCRIPTION = "0" AND SHOPCODE = "0" AND GLASSTYPEID ="0" Then
		response.write ""
	Else 
		If UPDATEFLAG = 1 Then
			Set rs6 = Server.CreateObject("adodb.recordset") 
			strSQL6 = "UPDATE XQSU_GlassTypes SET Type='" & GLASSTYPE& "', Description='" & DESCRIPTION & "', ShopCode='" & SHOPCODE & "', Status='" & STATUS & "'  WHERE ID = " & GLASSTYPEID
			rs6.Cursortype = 2
			rs6.Locktype = 3
			rs6.Open strSQL6, DBConnection

			If gi_Mode = c_MODE_ACCESS Or gi_Mode = c_MODE_SQL_SERVER Then
				GLASSTYPE = "0"
				DESCRIPTION = "0"
				SHOPCODE = "0"
				STATUS = "0"
				GLASSTYPEID = "0"
				UPDATEFLAG = "0"
			End If
		End if
	End if

DbCloseAll

End Function

	If  GLASSTYPE = "0" AND DESCRIPTION = "0" AND GLASSTYPEID ="0" Then
		response.write ""
	Else
%>
			<fieldset>

				<!--Entry form to edit existing Type and Description and then replace the old one in the database-->

			<div class="row">
				<% response.write"<input type='hidden' name='ID' id='GLASSTYPEID' value = '"& GLASSTYPEID & "' readonly = 'readonly'> "
				%>
			<div>
			<div class="row">
				<% response.write"<input type='hidden' name='UPDATE' id='UPDATE' value = '"& FLAG & "' readonly = 'readonly'> "
				%>
			<div>
			<div class="row">
				<label> Type: </label>
				<% response.write "<input type='text' name='GlassType' value = '" & GLASSTYPE & "' id='GLASSTYPE' > "
				%>
			</div>

			<div class="row">
				<label> Description: </label>
				<%response.write "<input type='text' name='Description' value = '" & DESCRIPTION & "' id='DESCRIPTION' >"
				%>
			</div>
			<div class="row">
				<label> Shop Code: </label>
				<%response.write "<input type='text' name='Shopcode' value = '" & ShopCode & "' id='ShopCode' >"
				%>
			</div>
			<div class="row">
				<label> Status: </label>
				<%response.write "<input type='text' name='Status' value = '" & STATUS & "' id='Status' >"
				%>
			</div>

			<a class="whiteButton" href="javascript:managetypes.submit()">Submit</a>

			<!-- SUBMIT BUTTON has to change to Edit the existing field rather than create new-->
</fieldset>
<%
		end if
%>
<ul id="Profiles" title="Glass" selected="true">
<li><table border ='1'>
<tr><th>Optima Glass Type</th><th>Description</th><th>ShopCode</th><th>Status</th></tr>

<%

	Set DBConnection = Server.CreateObject("adodb.connection")
	DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp
	DBConnection.Open DSN

	Set rs5 = Server.CreateObject("adodb.recordset")
	strSQ5L = "Select * FROM XQSU_GlassTypes ORDER BY TYPE ASC"
	rs5.Cursortype = 2
	rs5.Locktype = 3
	rs5.Open strSQ5L, DBConnection

do while not rs5.eof
	response.write "<tr>"
	response.write "<td><a href='glasstypesManage.asp?GlassType=" & rs5("Type") & "&Description=" & Server.URLEncode(rs5("Description") & "") & " &ShopCode=" & rs5("Shopcode") & "&ID=" & rs5("ID") & "&Status=" & rs5("Status") & "&UDPATE=0' target='_self'>" & rs5("Type") & "</a></td>"
	response.write"<td>" & rs5("Description") & "</td>"
	response.write"<td>" & rs5("ShopCode") & "</td>"
	response.write"<td>" & rs5("Status") & "</td>"
	response.write "</tr>"
	rs5.movenext
loop

%>
	</table></li>
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
