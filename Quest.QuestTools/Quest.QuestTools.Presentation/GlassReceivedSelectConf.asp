<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created Janaury 28th, 2015 - by Michael Bernholtz --> 
<!--Add Received Confirmation page - from GlassReceivedSelect.asp Requested by Joe De Francesco-->
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
                <a class="button leftButton" type="cancel" href="GlassReceivedSelect.asp" target="_self">Select</a>
    </div>

<form id="conf" title="Glass Edited" class="panel" name="conf" action= "GlassReceivedSelect.asp" method="GET" target="_self" selected="true" >              

        <h2>Stock Edited</h2>

<%
	GIDList = ""
	Received = Request.QueryString("Received")

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

	For Each item In Request.QueryString("GID")
		GID = item
		GIDList = GIDList & GID & ", " 	

		'Set Glass Inventory Update Statement
		StrSQL3 = "UPDATE Z_GLASSDB  SET [ExtRECEIVED]= '" & RECEIVED & "',[IntRECEIVED]= '" & RECEIVED & "' WHERE ID = " & GID 
		'Get a Record Set
		Set RS3 = DBConnection.Execute(strSQL3)
	Next

	DbCloseAll

End Function

%>

<ul id="Report" title="Added" selected="true">

<%

		Response.Write "<li>Records Received:</li>"
		Response.Write "<li>  Updated to : " & GIDList & "</li>"

%>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Received by Select</a>

            </form>

</body>
</html>

<%
'DBConnection.close
'set DBConnection=nothing
%>

