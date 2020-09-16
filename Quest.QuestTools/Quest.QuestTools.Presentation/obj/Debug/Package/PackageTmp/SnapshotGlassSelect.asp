<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_secondary.asp"-->
<!-- Created for Mahesh Mohanlall and Vanessa Abraham by Michael Bernholtz April 2019-->
<!-- Read and Report Glass information -->
<!-- SnapshotGlassSelect SnapshotGlassValue-->
<!-- Date: September 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add glass pricing to Snapshot Glass page.
-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Snapshot Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	<script type="text/javascript">
    function showForm() {
			
			var Country= document.getElementById("Country").value;
			if (Country == "USA") {
				document.getElementById("reportname1").style.display = "block";
				document.getElementById("reportname2").style.display = "none";
			} else {
				document.getElementById("reportname1").style.display = "none";
				document.getElementById("reportname2").style.display = "block";
			}
			
		}
		</script>
			
    </head>
<body onload = "showForm()">
  <div class="toolbar">
        <h1 id="pageTitle">Glass Snapshot Report</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite	%>#_Inv" target="_self">Stock<%response.write HomeSiteSuffix%></a>
				<div style=""><a class="button leftButton" type="cancel" href="SnapShotGlassSelect.asp" target="_self" style="left: 70px;">Inventory Report</a></div>
				<div style=""><a class="button leftButton" type="cancel" href="InventoryReportGlassPrices.asp" target="_self" style="left: 195px;">Glass Pricing</a></div>
    </div>

    <form id="snapshot" title="Take Inventory Snapshot" class="panel" name="snapshot" action="SnapShotGlassValue.asp" method="GET" target="_self" selected="true">
        <h2>Select a Snapshot of Glass Inventory</h2>
        <fieldset>

			<div class="row">
				<label>Report Country (USA / CANADA)</label>
				<select id='Country' name="Country" onchange="showForm()" >
		<% 
		if CountryLocation = "USA" then 
		else
		%>
					<Option value ='CANADA'>CANADA</Option>
		<%
		end if
		%>
					<Option value ='USA'>USA</Option>
				</select>
            </div>
			<div class="row"  id='reportname1' style="display:none">
			<label>Inventory Period (USA)</label>
			<select id='reportname' name="reportname1" >
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT ReportName FROM INV_Reports WHERE ReportName LIKE '%QC_Glass_USA%' ORDER BY ID DESC"
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
rs.close
set rs=nothing
%>	
			</select>
			</div>
			
			<div class="row"  id='reportname2' style="display:none">
			<label>Inventory Period (CANADA)</label>
			<select id='reportname' name="reportname2" >
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT ReportName FROM INV_Reports WHERE ReportName LIKE '%QC_Glass0%' or  ReportName LIKE '%QC_Glass1%' ORDER BY ID DESC"
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
rs.close
set rs=nothing
%>	
			</select>
            </div>


		
		
		</fieldset>
        <BR>
        <a class="whiteButton" href="javascript:snapshot.submit()">Submit</a>  
            </form>

<%

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
%>
	 
</body>
</html>
