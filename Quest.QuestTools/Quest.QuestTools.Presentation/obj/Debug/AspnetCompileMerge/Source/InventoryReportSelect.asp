<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Created at Request of Shaun Levy with permission from Jody Cash -->
<!--Input form to Collect Aluminium Price and choose a month to get valuation-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Inventory Report</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

    <form id="snapshot" title="Take Inventory Snapshot" class="panel" name="snapshot" action="InventoryReportValue.asp" method="GET" target="_self" selected="true">
        <h2>Enter Aluminium Price and Select a Snapshot of Inventory</h2>
        <fieldset>

	        <div class="row">
                <label>Price ($)</label>
                <input type="number" name='AlumPrice' id='AlumPrice' value ="3.95" step="any" min = 0  >
        </div>
			<div class="row">
			<label>Inventory Period</label>
			<select id='reportname' name="reportname" >
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT ReportName FROM INV_Reports WHERE ReportName LIKE '%Y_INV%' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection2

Do while not rs.eof

	Response.write "<Option value ='"
	Response.write rs.fields("ReportName")
	Response.write "'>" 
	Response.write rs.fields("ReportName")
	Response.write "</option>"
	

rs.movenext
loop
%>	
			</select>
            </div>
		
		
		</fieldset>
        <BR>
        <a class="whiteButton" href="javascript:snapshot.submit()">Full Report</a>   
		<input type="submit" value = "Sheets" class="whiteButton" onclick="snapshot.action='InventoryReportValueSheet.asp';"></input>
		<input type="submit" value = "Plastic" class="whiteButton" onclick="snapshot.action='InventoryReportValuePlastic.asp';"></input>			
        <a class="redButton" href="inventoryReportCreate.asp">Create New Inventory Snapshot</a>   
		
	
            </form>
         
<%
rs.close
set rs=nothing
DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
%>
	 
</body>
</html>
