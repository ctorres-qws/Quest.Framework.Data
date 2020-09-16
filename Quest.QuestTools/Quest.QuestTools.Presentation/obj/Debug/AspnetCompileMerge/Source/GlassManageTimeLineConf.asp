<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Edit Confirmation Page for Glass Items-->
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

<%

Gid = request.querystring("Gid")
ticket = Request.QueryString("ticket")

select case ticket
	case "active"
		returnEmail = "GlassToolViewActive.asp"
	case "optima"
		returnEmail = "GlassToolViewOptima.asp"
	case "waiting"
		returnEmail = "GlassToolViewWait.asp"
	case "received"
		returnEmail = "GlassToolViewReceived.asp"
	case "completed"
		returnEmail = "GlassToolViewCompleted.asp"
	case "shipped"
		returnEmail = "GlassToolViewShipped.asp"
	case Else
		returnEmail = "GlassManage.asp"
	End select
%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassManageForm.asp?GID=<% response.write Gid %>" target="_self">Edit Glass</a>
    </div>
    
      
    
<form id="conf" title="Glass Edited" class="panel" name="conf" action="<%response.write returnEmail%>" method="GET" target="_self" selected="true" >

        <h2>Stock Edited</h2>

<%

BARCODE = "GT" & ID

REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
INPUTDATE = REQUEST.QueryString("INPUTDATE")
OPTIMADATE = REQUEST.QueryString("OPTIMADATE")
EXTEXPECTED = REQUEST.QueryString("EXTEXPECTED")
EXTRECEIVED = REQUEST.QueryString("EXTRECEIVED")
INTEXPECTED = REQUEST.QueryString("INTEXPECTED")
INTERECEIVED = REQUEST.QueryString("INTRECEIVED")
COMPLETEDDATE = REQUEST.QueryString("COMPLETEDDATE")
SHIPDATE = REQUEST.QueryString("SHIPDATE")
CARDINALSENT = REQUEST.QueryString("CARDNIALSENT")
CARDINALEXPECTED = REQUEST.QueryString("CARDINALEXPECTED")
CARDINALRECEIVED = REQUEST.QueryString("CARDINALRECEIVED")
QUICKTEMPSENT = REQUEST.QueryString("QUICKTEMPSENT")
QUICKTEMPRECEIVED = REQUEST.QueryString("QUICKTEMPRECEIVED")

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
	StrSQL = "UPDATE Z_GLASSDB  SET [REQUIREDDATE]= '" & REQUIREDDATE & "',[INPUTDATE]= '" & INPUTDATE & "',[COMPLETEDDATE]= '" & COMPLETEDDATE & "',[SHIPDATE]= '" & SHIPDATE & "',[CARDINALSENT]= '" & CARDNINALSENT & "',[CARDINALEXPECTED]= '" & CARDINALEXPECTED & "',[CARDINALRECEIVED]= '" & CARDINALRECEIVED & "',[QUICKTEMPSENT]= '" & QUICKTEMPSENT & "',[QUICKTEMPRECEIVED]= '" & QUICKTEMPRECEIVED & "',[OPTIMADATE]= '" & OPTIMADATE & "',[EXTEXPECTED]= '" & EXTEXPECTED & "',[INTEXPECTED]= '" & INTEXPECTED & "',[EXTRECEIVED]= '" & EXTRECEIVED & "',[INTRECEIVED]= '" & INTRECEIVED & "'  WHERE ID = " & GID

	'Get a Record Set
	Set RS = DBConnection.Execute(strSQL)

	DbCloseAll

End Function

%>

<ul id="Report" title="Added" selected="true">

<%

		Response.Write "<li>Inventory GLASS Time Line Edited:</li>"
		Response.Write "<li> Required Date: " & REQUIREDDATE & "</li>"
		Response.Write "<li> Optima Date: " & OPTIMADATE & "</li>"
		Response.Write "<li> Exterior Glass Expected Date: " & EXTEXPECTED & "</li>"
		Response.Write "<li> Exterior Glass Received Date: " & EXTRECEIVED & "</li>"
		Response.Write "<li> Interior Glass Expected Date: " & INTEXPECTED & "</li>"
		Response.Write "<li> Interior Glass Received Date: " & INTRECEIVED & "</li>"
		Response.Write "<li> Unit Completed Date: " & COMPLETEDDATE & "</li>"
		Response.Write "<li> Unit Shipped Date: " & SHIPDATE & "</li>"
		Response.Write "<li> Date Sent to Cardinal: " & CARDINALSENT & "</li>"
		Response.Write "<li> Date Expected back from Cardinal: " & CARDINALEXPECTED & "</li>"
		Response.Write "<li> Date Unit Received from Cardinal: " & CARDINALRECEIVED & "</li>"
		Response.Write "<li> Date Sent to Quick Temp: " & QUICKTEMPSENT & "</li>"
		Response.Write "<li> Date Unit Received from QuickTemp: " & QUICKTEMPRECEIVED & "</li>"

%>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Home</a>

            </form>

</body>
</html>

<%

'DBConnection.close
'set DBConnection=nothing
%>

