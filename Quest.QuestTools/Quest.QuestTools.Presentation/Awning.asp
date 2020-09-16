<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<%
Server.ScriptTimeout=2000
%>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

        <ul id="Profiles" title=" Glass Report - All Active" selected="true">

		<%
		
		OVCOUNTER = 0

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT ShipDate,Job,Floor,tag FROM X_SHIPPING WHERE SHIPDATE > #2018-04-01# AND SHIPDATE < #2018-06-01#  AND JOB <> '' AND FLOOR <> '' AND TAG <> ''  ORDER BY JOB DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM Styles ORDER BY ID DESC"
rs3.Cursortype = GetDBCursorType
rs3.Locktype = GetDBLockType
rs3.Open strSQL3, DBConnection

DO while not rs.eof
idcounter = idcounter + 1
Response.write idcounter & "   -   "

Set rs2 = Server.CreateObject("adodb.recordset")
Response.write RS("JOB")
Response.write RS("FLOOR")
Response.write RS("TAG")



strSQL2 = "SELECT Style, Job, FLoor, Tag FROM [" & rs("job") & "] WHERE JOB = '" & rs("job") & "' AND Floor = '" & rs("floor") & "' AND Tag = '-" & rs("tag") & "'  ORDER BY JOB DESC"

rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection

Response.write rs2("Style")

rs3.filter = "Name = '" & rs2("style") & "'"
if not rs3.eof then
	if rs3("O1") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
		
	end if
	if rs3("O2") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O3") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O4") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O5") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O6") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O7") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
	if rs3("O8") = "OV" then
		OVCOUNTER = OVCOUNTER + 1
	end if
Response.write " - " & OVCOUNTER
	end if


rs2.close
set rs2 = nothing
Response.Write "&nbsp;<br /><img src='images/done.gif'>"


rs.movenext
loop

%>
		
		
		
		<% 


response.write "<li class='group'>" & OVCOUNTER & " </li>"

rs.close
set rs = nothing
rs3.close
set rs3 = nothing
DBConnection.close 
set DBConnection = nothing

%>
      </ul>
</body>
</html>
