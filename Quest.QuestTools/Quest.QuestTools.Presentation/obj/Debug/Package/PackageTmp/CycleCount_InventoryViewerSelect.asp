<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Date: January 13, 2019
	Modified By: Michelle Dungo
	Changes: Modified to generate cycle count viewer for sheets and hardware
	
	Date: January 17, 2019
	Modified By: Michelle Dungo
	Changes: Modified to generate cycle count viewer for plastic
-->
<!-- CycleCount_InventoryViewerSelect CycleCount_InventoryViewer[Sheets|Hardware|Plastic] -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Cycle Count</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
	<script>	
		function viewHardwareReport() {
			
			cyclecount.action = "CycleCount_InventoryViewerHardware.asp";	
			cyclecount.submit();
		}

		function viewSheetReport() {
			
			cyclecount.action = "CycleCount_InventoryViewerSheets.asp";
			cyclecount.submit();
		}	
		
		function viewPlasticReport() {
			
			cyclecount.action = "CycleCount_InventoryViewerPlastic.asp";
			cyclecount.submit();
		}			
	</script>		
    </head>
<body>
  <div class="toolbar">
        <h1 id="pageTitle">Cycle Count Viewer</h1>
			<a class="button leftButton" type="cancel" href="index.html#_TmpINV" target="_self">INV-CC</a>
    </div>

    <form id="cyclecount" title="Cycle Count Viewer" class="panel" name="cyclecount" action="" method="GET" target="_self" selected="true">
        <h2>Select Modified Date Range</h2>
        <fieldset>

			<div class="row">
			<label>Start Date</label>
				<input type="date" id="startDate" name="startDate" required>
            </div>
			<div class="row">
				<label>End Date</label>
					<input type="date" id="endDate" name="endDate" required>
            </div>
		</fieldset>
        <BR>
        <a class="whiteButton" href="javascript: viewHardwareReport()">Hardware Adjustment Report</a>  
		<a class="whiteButton" href="javascript: viewSheetReport();">Sheet Adjustment Report</a>		
		<a class="whiteButton" href="javascript: viewPlasticReport();">Plastic Adjustment Report</a>				
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
