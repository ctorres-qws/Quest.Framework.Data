<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_Quest_InventoryReports.asp"-->
		 
<!-- Stored Procedure to be run on the first of each month-->
<!-- Removes all the old CUT, HCUT, DMSAW Tables by matching 2 month old data from Stored Procedure-->
<!-- 


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
if currentMonth > 3 then 
	TwoAgoMonth = currentMonth -2
	TwoAgoYear = currentYear
end if
				' If it is February, 2 Months Ago is December Last Year
if currentMonth =2 then 
	TwoAgoMonth = 12
	TwoAgoYear = currentYear- 1
end if
				' If it is January, 2 Months Ago is November Last Year
if currentMonth =1 then 
	TwoAgoMonth = 11
	TwoAgoYear = currentYear- 1
end if

				' Adds a 0 to Num 1-9 for consistency 
if TwoAgoMonth < 10 then
	TwoAgoMonth = "0" & TwoAgoMonth
end if
 
 





cycle = "c1"
i =1

do While i <10000
	

	i =i+1
	CutTable = "Cut_" & RecordJob & RecordFloor & cycle 
	
	Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "DROP TABLE " & CutTable & i &  " WHERE EXISTS ( SELECT * FROM  " & CutTable & i & ")"
		rs2.Open strSQL2, DBConnection2
		
		If rs2.State = 1 Then 
			rs2.Close
			set rs2 = nothing
		End If


loop
 
rs.close
set rs = nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
 
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
	<li>Items copied</li>
	<%end if%>
	</ul>
 

<% 

%>
 
</body>
</html>
