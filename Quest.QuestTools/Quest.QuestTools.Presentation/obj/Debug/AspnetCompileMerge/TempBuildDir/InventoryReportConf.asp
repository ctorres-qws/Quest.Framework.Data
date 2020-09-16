<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Created at Request of Shaun Levy with permission from Jody Cash -->
<!-- Multiple attempts were made to connect SQL Database to ACCESS DATABASE, all unsucessful, this is a patch - second ACCESS Database-->
<!-- Eventually use dbpath_QuestAccess_QuestSQL.asp -->

<!--Input form to take snapshot of Y_Inv and add to another ACCESS DATABASE using new DBPATH in next page-->

<!--Update April 2019 to include Glass from QualityControlDB for Glass Snapshot, Michael Bernholtz - only Canada-->
<!--Date: July 16, 2019
	Modified By: Michelle Dungo
	Changes: Modified to change field names for QC_GLASS and QC_GLASS_USA which don't exist
			Remove quotes when checking number values
			Force timeout for this page to prevent issues generating incomplete tables
-->

<%
Server.ScriptTimeout=4000
%>

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
</head>
<body>

<%
isError=FALSE
error = ""
Reporttime = Request.QueryString("snapmonth")
currentDate = Date()
Dim str_Where_I: str_Where_I = ""
Dim str_Where_H: str_Where_H = ""
Dim str_Where_G: str_Where_G = ""

' Sets the Month and Year of the Data to be Saved from Y_INV
Select Case Reporttime

Case "current"
	SnapMonth = Month(now)
	SnapYear = Year(now)
Case "previous"

	dt_Now = Now
	str_SnapDate = "1-" & ga_Months(Month(dt_Now)) & "-" & Year(dt_Now)
	str_Where_I = " WHERE DateIN < #" & str_SnapDate & "# "
	str_Where_H = " WHERE EnterDate < #" & str_SnapDate & "# "
	str_Where_G = " WHERE EntryDate < #" & str_SnapDate & "# "

	If Month(now) = 1 Then
		SnapMonth =  12
		SnapYear = Year(now) -1
	Else
		SnapMonth = Month(now)-1
		SnapYear = Year(now)
	End If
Case Else
	SnapMonth = Month(now)
	SnapYear = Year(now)
End Select

' Adds a 0 to Num 1-9 for consistency 
if SnapMonth < 10 then
	SnapMonth = "0" & SnapMonth
end if


 'Attempts to use SQL SERVER, DBConnection2  - [QWS-DEV].[dbo].Y_INV in QUESTACCESS_QUESTSQL
 ' Can successfully use both databases but cannot copy directly from one to the other.
'Removed the Case gi_mode code from Harj's Program and only run the Access

ProcessAccess(false)

Function ProcessAccess(isSQLServer)
	DbOpen DBConnection, isSQLServer
	DbOpenSecondary DBConnection2, isSQLServer

	ReportCreate = "Y_INV" & SnapMonth & SnapYear
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "Select * from INV_Reports"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection2
	rs.filter = "ReportName = '" & ReportCreate & "'"

	If rs.eof Then

		rs.addnew
		rs.fields("ReportName") = "Y_INV" & SnapMonth & SnapYear
		rs.fields("CreatedDate") = currentDate
		rs.fields("SnapMonth") = SnapMonth
		rs.fields("SnapYear") = SnapYear
		rs.update

		rs.addnew
		rs.fields("ReportName") = "X_Barcode" & SnapMonth & SnapYear
		rs.fields("CreatedDate") = currentDate
		rs.fields("SnapMonth") = SnapMonth
		rs.fields("SnapYear") = SnapYear
		rs.update

		rs.addnew
		rs.fields("ReportName") = "Y_Hardware" & SnapMonth & SnapYear
		rs.fields("CreatedDate") = currentDate
		rs.fields("SnapMonth") = SnapMonth
		rs.fields("SnapYear") = SnapYear
		rs.update
		
		rs.addnew
		rs.fields("ReportName") = "QC_Glass" & SnapMonth & SnapYear
		rs.fields("CreatedDate") = currentDate
		rs.fields("SnapMonth") = SnapMonth
		rs.fields("SnapYear") = SnapYear
		rs.update
		
		rs.addnew
		rs.fields("ReportName") = "QC_Glass_USA" & SnapMonth & SnapYear
		rs.fields("CreatedDate") = currentDate
		rs.fields("SnapMonth") = SnapMonth
		rs.fields("SnapYear") = SnapYear
		rs.update

		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "Select * into Y_INV" & SnapMonth & SnapYear & " FROM [MS Access;DATABASE=" & "F:\database\quest.mdb" & "].[Y_INV]"
		rs2.Open strSQL2, DBConnection2
		set rs2 = nothing

		Set rs5 = Server.CreateObject("adodb.recordset")
		strSQL5 = "Select * into Y_Hardware" & SnapMonth & SnapYear & " FROM [MS Access;DATABASE=" & "F:\database\quest.mdb" & "].[Y_Hardware]"
		rs5.Open strSQL5, DBConnection2
		set rs5 = nothing
		
		Set rs6 = Server.CreateObject("adodb.recordset")
		strSQL6 = "Select * into QC_Glass" & SnapMonth & SnapYear & " FROM [MS Access;DATABASE=" & "F:\database\QualityControlDB.mdb" & "].[QC_GLASS]"
		rs6.Open strSQL6, DBConnection2
		set rs6 = nothing
	
		Set rs9 = Server.CreateObject("adodb.recordset")
		strSQL9 = "Select * into QC_Glass_USA" & SnapMonth & SnapYear & " FROM [MS Access;DATABASE=" & "F:\database\QualityControlDB.mdb" & "].[QC_GLASS_USA]"
		rs9.Open strSQL9, DBConnection2
		set rs9 = nothing

		Set rs3 = Server.CreateObject("adodb.recordset")
		strSQL3 = "Select * into X_Barcode" & SnapMonth & SnapYear & " FROM [MS Access;DATABASE=" & "F:\database\quest.mdb" & "].[X_BARCODE] WHERE MONTH = " & SnapMonth & "AND YEAR = " & SnapYear 
		rs3.Open strSQL3, DBConnection2
		set rs3 = nothing

		'Added a Component to save the KGM along side the value
		Set rs4 = Server.CreateObject("adodb.recordset")
		strSQL4 = "Select * FROM Y_INV" & SnapMonth & SnapYear
		rs4.Cursortype = 2
		rs4.Locktype = 3
		rs4.Open strSQL4, DBConnection2

		Do While Not rs4.eof
			invpart = rs4("part")
			'Create a Query
			SQL5 = "SELECT * FROM Y_MASTER where PART = '" & invpart & "' order BY ID DESC"
			'Get a Record Set
			Set RS5 = DBConnection.Execute(SQL5)

			If RS5.EOF Then
			Else
				rs4("kgm") = FormatNumber(rs5("kgm"),4)
				rs4.update
			End If

			RS5.close
			Set RS5 = nothing

			rs4.movenext
		Loop
		
		'Added a Component to save the PriceSQFT along side the value of Glass
		Set rs7 = Server.CreateObject("adodb.recordset")
		strSQL7 = "Select * FROM QC_GLASS" & SnapMonth & SnapYear
		rs7.Cursortype = 2
		rs7.Locktype = 3
		rs7.Open strSQL7, DBConnection2

		Do While Not rs7.eof
			invpart = rs7("MasterID")
			'Create a Query
			SQL8 = "SELECT * FROM [MS Access;DATABASE=" & "F:\database\QualityControlDB.mdb" & "].[QC_MASTER_GLASS] where ID = " & invpart & " order BY ID DESC"
			'Get a Record Set
			Set rs8 = DBConnection.Execute(SQL8)

			If rs8.EOF Then
			Else
				rs7("PriceSQFT") = FormatNumber(rs8("Price"),4)
				rs7.update
			End If

			rs8.close
			Set rs8 = nothing

		rs7.movenext
		Loop
			rs7.close
			Set rs7 = nothing
			
			
		'Added a Component to save the PriceSQFT along side the value of Glass USA
		Set rs7 = Server.CreateObject("adodb.recordset")
		strSQL7 = "Select * FROM QC_GLASS_USA" & SnapMonth & SnapYear
		rs7.Cursortype = 2
		rs7.Locktype = 3
		rs7.Open strSQL7, DBConnection2

		Do While Not rs7.eof
			invpart = rs7("MasterID")
			'Create a Query
			SQL8 = "SELECT * FROM [MS Access;DATABASE=" & "F:\database\QualityControlDB.mdb" & "].[QC_MASTER_GLASS] where ID = " & invpart & " order BY ID DESC"
			'Get a Record Set
			Set rs8 = DBConnection.Execute(SQL8)

			If rs8.EOF Then
			Else
				rs7("PriceSQFT") = FormatNumber(rs8("Price"),4)
				rs7.update
			End If

			rs8.close
			Set rs8 = nothing

		rs7.movenext
		Loop
			rs7.close
			Set rs7 = nothing

	Else

		IsError = true
		error = "Y_INV" & SnapMonth & SnapYear & ": Already created"

	End If

	DbCloseAllAndSecondary

End Function

%>

	<div class="toolbar">
        <h1 id="pageTitle">Monthly Snapshot</h1>
		<% 
			if CountryLocation = "USA" then 
				BackButton = "indexTexas.html#_Inv"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Inv"
				HomeSiteSuffix = ""
			end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Inv <%response.write HomeSiteSuffix%></a>
    </div>
	

	<ul id="Profiles" title="SnapShot of Inventory " selected="true">
<%
	If IsError = True Then
%>
		<li>Report Not Generated: <%response.write error %>
<%
	Else
%>
		<li>Inventory Report Generated: <%response.write SnapMonth & "/" & SnapYear %>
		<li>Monthly Barcode Backup Created: <%response.write SnapMonth & "/" & SnapYear %>
		<li>Hardware Inventory Report Generated: <%response.write SnapMonth & "/" & SnapYear %>
		<li>Glass Inventory Report Generated: <%response.write SnapMonth & "/" & SnapYear %>
		<li><a href="InventoryReportSelect.asp" target="_self"><b>GO TO</b> Inventory Report</a></li>
<%
	End If
%>
	</ul>

<%

'rs.close
'set rs = nothing

'DBConnection.close
'set DBConnection = nothing
'DBConnection2.close
'set DBConnection2 = nothing

%>

</body>
</html>
