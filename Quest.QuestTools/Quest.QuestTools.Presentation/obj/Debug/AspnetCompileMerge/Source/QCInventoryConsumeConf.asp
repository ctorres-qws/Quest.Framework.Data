<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 26th, by Michael Bernholtz - Consumed Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant Misc go to QC_Misc-->
<!-- Consumed Items are given a Consumed Date and Quantity (Glass only) set to 0-->
<!-- February 2019 - Added USA Tables to database-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Consume QC Inventory</title>
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
InventoryType = REQUEST.QueryString("InventoryType")
qcid = request.querystring("qcid")

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
                <a <a class="button leftButton" type="cancel" href="QCInventoryEditForm.asp?InventoryType=<% response.write InventoryType %>&QCID=<% response.write qcid %>" target="_self">Edit Stock</a>

    </div>

<%
ConsumeDate = Date()

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

	DbOpenQC DBConnection, isSQLServer

	Select Case InventoryType
	Case "QCGlass"
			'Set Glass INVENTORY Consume Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_GLASS_USA SET Quantity= 0 , ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_GLASS SET Quantity= 0 , ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	Case "QCSpacer"
			'Set SPACER INVENTORY Consume Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_SPACER_USA SET ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_SPACER SET ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	Case "QCSealant"
			'Set SEALANT INVENTORY Consume Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_SEALANT_USA SET ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_SEALANT SET ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			end if
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	Case "QCMisc"
			'Set Glass INVENTORY Consume Statement
			if CountryLocation = "USA" then
				StrSQL = "UPDATE QC_Misc_USA SET Quantity= 0 , ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			else
				StrSQL = "UPDATE QC_Misc SET Quantity= 0 , ConsumeDate='" & ConsumeDate & "'  WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	End Select

	DbCloseAll

End Function

		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>

<form id="conf" title="Consume Stock" class="panel" name="conf" action="<%response.write Homesite%>#_QC" method="GET" target="_self" selected="true" >              

        <h2>Inventory Item Consumed</h2>
		<div class="row">

		</div>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
            
            </form>
<%
On Error Resume Next
DBConnection.close
set DBConnection=nothing
%>
			
</body>
</html>



