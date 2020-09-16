<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created August 12th, by Michael Bernholtz - Delete Confirmation for items in Optimization Log-->
<!-- Optimization Log Table created for Victor, Implemented by Michael Bernholtz-->  

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Delete Log Item</title>
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
OPID = request.querystring("OPID")

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a <a class="button leftButton" type="cancel" href="OptimizationLogManage.asp%>" target="_self">Manage</a>

    </div>
    
<%

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

		'Set Optimization Log Delete Statement
			StrSQL = "DELETE FROM OptimizeLog WHERE ID = " & OPID
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)

DbCloseAll

End Function

%>
    
<form id="conf" title="Delete Stock" class="panel" name="conf" action="index.html#_GlassP" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Deleted</h2>
		<div class="row">

		</div>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Back to Glass Production</a>
            
            </form>

            
    
</body>
</html>

<% 


DBConnection.close
set DBConnection=nothing
%>

