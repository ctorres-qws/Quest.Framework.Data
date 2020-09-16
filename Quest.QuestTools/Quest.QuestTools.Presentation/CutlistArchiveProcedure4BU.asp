<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_Quest_ArchiveLists.asp"-->
		 
<!-- Stored Procedure to be run on the first of each month-->
<!-- Deletes  all the QSU, QSP, PANEL Tables -->
<!-- Updated August 2014, To mark Cutlists inactive from Z_Cutlists -->

<!-- Works, but code is ugly This shoudl be cleaned up -->
<!-- Currently works manually only by Table name selected -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Archive Program</title>
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
        <h1 id="pageTitle">Delete QSU/QSP/PANEL</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>
   
   
   
    <ul id="Profiles" title="SnapShot of Inventory " selected="true">
 
 
 <%
 


 
'Collect TableNames from Schema Table 
Const adSchemaTables = 20
Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")


 TableCount = 0
 TableFull = 0
 TableDeleted = 0

i = 1
   
Do Until i >= 4
   
	Select Case i
		Case 1
			TableName = "QSU_*"
		Case 2
			TableName = "QSP_*"
		Case 3
			TableName = "PANEL_*"

	End Select

   
	rs.filter = "TABLE_NAME LIKE '" & TableName & "' "
 
	Do while not rs.eof
		TableCount = TableCount + 1
		TableName = rs("TABLE_NAME")
		TableCheckStatus = FALSE
	
		Set Tablecheck = Server.CreateObject("adodb.recordset")
		TC_SQL = "SELECT * From [" & TableName & "]"
		Tablecheck.Cursortype = 1
		Tablecheck.Locktype = 3
		Tablecheck.Open TC_SQL, DBConnection
	
		StatusDone = 0
		StatusCount = 0
	
		if Tablecheck.RecordCount > 0 then
				TableCheckstatus = TRUE
		end if

		Tablecheck.close
		Set Tablecheck = nothing
		

		if TableCheckstatus = FALSE then
		else
		TableFull = TableFull + 1

	'		Set rs2 = Server.CreateObject("adodb.recordset")
	'		strSQL2 = "Select * into [MS Access;DATABASE=f:\database\ArchiveLists.mdb]." & TableName &  " FROM [MS Access;DATABASE=f:\database\quest.mdb]." & TableName
	'			On Error Resume Next  
	'		rs2.Open strSQL2, DBConnection2
	'			On Error GoTo 0
				
		
			SQL3 = "Drop TABLE " & TableName 
				On Error Resume Next  
			set RS3 = DBConnection.Execute(SQL3)
			if Err.Number = 0 then
				TableDeleted = TableDeleted + 1
			end if 

				On Error GoTo 0
			
			SQL4 = "UPDATE Z_Cutlists SET ACTIVE = FALSE WHERE CUTLIST = '" &TableName &  "'"
				On Error Resume Next  
			RS4 = DBConnection.Execute(SQL4)
				On Error GoTo 0

		end if
		
	rs.movenext	
	Loop
	

%>
	<li><B><U><%Response.write TableName %></U></B></li>
	<li>Tables Counted: <%response.write TableCount %></li>
	<li>Tables Full: <%response.write TableFull %></li>
	<li>Tables Deleted: <%response.write TableDeleted %></li>
	<li>Tables Remaining: <%response.write TableCount - TableDeleted %></li>

	</ul>
	
 <%
 
i = i+ 1
Loop

rs.close
set rs= nothing



DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing
 
%>
 
</body>
</html>
