<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_Quest_ArchiveLists.asp"-->
		 
<!-- Stored Procedure to be run on the first of each month-->
<!-- Removes all the old CUT, HCUT, DMSAW Tables by matching 2 month old data from Stored Procedure-->
<!-- Updated August 2014, To mark Cutlists inactive from Z_Cutlists -->


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
 

 TableCount = 0
 

	
	TableName = "Z_RATES"

	Set rs2 = Server.CreateObject("adodb.recordset")
Response.write "1"
		strSQL2 = "Select * into [MS Access;DATABASE=f:\database\Archive2.mdb]." & TableName &  " FROM [MS Access;DATABASE=f:\database\quest.mdb]." & TableName
			On Error Resume Next  
				rs2.Open strSQL2, DBConnection2
			On Error GoTo 0
			
		
	'		SQL3 = "Drop TABLE " & TableName 
	'			On Error Resume Next  
	'		Set RS3 = DBConnection.Execute(SQL3)
	'			On Error GoTo 0
		

Response.write "2"
		If rs2.State = 1 Then 
	Response.write "3"
			TableCount = TableCount + 1
			rs2.Close
			set rs2 = nothing
		End If
		




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
	<li>Items copied: <%response.write TableCount %></li>
	<%response.write Copied %>
	<%end if%>
	</ul>
 

<% 

%>
 
</body>
</html>
