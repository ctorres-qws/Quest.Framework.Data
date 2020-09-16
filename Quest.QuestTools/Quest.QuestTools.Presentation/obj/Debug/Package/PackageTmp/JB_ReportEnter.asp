<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Jamb Receptor View Form Entry, Choose by Job and Floor - Checks JB_XXX##-->
<!-- Basic Entry form here to get Job And Floor JB_ReportEnter.asp sends to JB_Report.asp-->
<!-- JB_Report Designed by Ariel Aziza, Coded by Michael Bernholtz, October 2018 -->


<!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
  <title>Jamb Receptor Report</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_JB" target="_self">Jamb Receptor</a>
    </div>

    <form id="enter" title="Jamb Receptor" class="panel" name="enter" action="JB_Report.asp" method="GET" target="_self" selected="true">
              
	
	<h2>Choose Job and Floor for Jamb Receptor Report</h2>
		<fieldset>	
<div class="row">
				<label>JOB</Label>
				<select name="Job">
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT JOB FROM Z_JOBS ORDER BY JOB ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.movefirst
Do While Not rs.eof

Response.Write "<option value='"
Response.Write rs("JOB")
Response.Write "'>"
Response.Write rs("JOB")
Response.write "</option>"

rs.movenext

loop
rs.close
set rs=nothing
%></select>
            </div>
		
		<div class="row">
			<label>Floor</label>
			<input type="text" name='Floor' id='Floor' >
		</div>
	</fieldset>
	<BR>
	<a class="whiteButton" href="javascript:enter.submit()">Submit</a>

</form>    

<%
DBConnection.close
set DBConnection = nothing
%>
</body>
</html>
