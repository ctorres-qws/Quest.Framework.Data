<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- Updated February 26th to include Consumed and the ability to clear consumption and reactivate -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Edit QC Inventory</title>
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

BackTag = REQUEST.QueryString("BackTag")
OPID = REQUEST.QueryString("Opid")

%>

	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        </div>
       
    
<form id="conf" title="Production" class="panel" name="conf" action="index.html#_GlassP" method="GET" target="_self" selected="true" >              

  
   
        <h2>This Glass Marked Back order: please close this tab to continue</h2>
  
<%       
CurrentDate = Date & "-" & Backtag
	if Backtag = "Ext" then
		'Set Sealant Inventory Update Statement
			StrSQL = "UPDATE Z_GlassDB SET backorderflag = '" & currentDate & "' , ExtReceived = NULL WHERE ID = " & OPID
		'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	end if
	if Backtag = "Int" then 
		'Set Sealant Inventory Update Statement
			StrSQL = "UPDATE Z_GlassDB SET backorderflag = '" & currentDate & "' , IntReceived = NULL WHERE ID = " & OPID
		'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	end if
%>				

<ul id="Report" title="Added" selected="true">
	
<%
		Response.Write "<li>Glass Marked on Backorder:</li>"
		Response.Write "<li> Received Date Removed for ID: " & OPID & "</li>"
		Response.Write "<li> BackOrderFlag: " & Date() & "</li>"
	

%>

        <BR>
       
         <a class="whiteButton" href="OptimizationLogManage.asp" target="_self"> Back</a>

            </form>

</body>
</html>

<% 

DBConnection.close
set DBConnection=nothing
%>

