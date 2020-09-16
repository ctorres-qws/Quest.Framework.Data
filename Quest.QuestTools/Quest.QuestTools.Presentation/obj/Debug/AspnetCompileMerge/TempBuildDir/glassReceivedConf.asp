<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 2015 - by Michael Bernholtz --> 
<!--Add RECEIVED Date Confirmation page - from GlassReceived.asp Requested by Joe-->
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
                <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tool </a>
    </div>

<form id="conf" title="Glass Edited" class="panel" name="conf" action= "GlassReceived.asp" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%

WORKORDER = REQUEST.QueryString("WORKORDER")
RECEIVED = Request.QueryString("RECEIVED")

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
	StrSQL1 = "UPDATE Z_GLASSDB  SET [ExtReceived]= '" & RECEIVED & "' WHERE ExtOrderNum LIKE '%" & WORKORDER & "%' "

	'Get a Record Set
	Set RS1 = DBConnection.Execute(strSQL1)
	StrSQL2 = "UPDATE Z_GLASSDB  SET [IntReceived]= '" & RECEIVED & "' WHERE IntOrderNum LIKE '%" & WORKORDER & "%' "

	'Get a Record Set
	Set RS2 = DBConnection.Execute(strSQL2)
	StrSQL3 = "UPDATE Z_GLASSDB  SET [ExtReceived]= '" & RECEIVED & "',[IntReceived]= '" & RECEIVED & "' WHERE PO LIKE '%" & WORKORDER & "%' "

	'Get a Record Set
	Set RS3 = DBConnection.Execute(strSQL3)

	DbCloseAll

End Function

%>

<ul id="Report" title="Added" selected="true">

<%
		Response.Write "<li> Receive Date added:</li>"
		Response.Write "<li>All Work Order Items for: " & WORKORDER & "</li>"
		Response.Write "<li> Received Date: " & RECEIVED & "</li>"
%>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Add Another</a>
</form>

<%
'DBConnection.close
'set DBConnection=nothing
%>

</body>
</html>

