<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_Quest_InventoryReports.asp"-->
		 
<!-- Stored Procedure to be run on the first of each month-->
<!-- Creates a copy of Y_INV (Snapshot of current Database information-->
<!-- Creates a Backup of X_BARCODE data from 2 Months previous AND Deletes them from X_Barcode-->
<!-- Searches for old MACHINING tables, DELETE them FROM QUest.mdb and save them to a backup access Database-->


<!--Input form to take snapshot of Y_Inv and add to another ACCESS DATABASE using new DBPATH in next page-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Inventory Snapshot</title>
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
 
Reporttime = Request.QueryString("snapmonth")
currentDate = Date()

currentMonth = Month(Now)
currentYear = Year(Now)

	' Set 2 Months Ago by Subtracting Two Months
TwoAgoMonth = Month(DateAdd("m",-2,Now))
TwoAgoYear = Year(DateAdd("m",-2,Now))

	' Adds a 0 to Num 1-9 for consistency 
if TwoAgoMonth < 10 then
	TwoAgoMonth = "0" & TwoAgoMonth
end if
 
 'Attempts to use SQL SERVER, DBConnection2  - [QWS-DEV].[dbo].Y_INV in QUESTACCESS_QUESTSQL
 ' Can successfully use both databases but cannot copy directly from one to the other.



Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * into [MS Access;DATABASE=f:\database\InventoryReports.mdb].[X_Barcode" & TwoAgoMonth & TwoAgoYear & "Arch] FROM [MS Access;DATABASE=f:\database\quest.mdb].[X_BARCODE] WHERE MONTH = " & TwoAgoMonth & "AND YEAR = " & TwoAgoYear 
rs.Open strSQL3, DBConnection2

%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
   
   
   
    <ul id="Profiles" title="SnapShot of Inventory " selected="true">
	<% if IsError = True then %>
	<li>Report Not Generated: <%response.write error %>
	<%else%>	
	<li>Monthly Barcode Backup Created: <%response.write TwoAgoMonth & "/" & TwoAgoYear %>	
	<%end if%>
	</ul>
 

<% 

rs.close
set rs = nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
%>
 
</body>
</html>
