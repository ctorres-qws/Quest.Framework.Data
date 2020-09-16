<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Recreated February 26th after deletion, by Michael Bernholtz - Reports shows Information from All QC Inventory Tables and uses matched information from the Master Tables-->
<!-- QC Inventory Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- One of 3 Tables - QC_GLASS, QC_SPACER, QC_SEALANT-->
<!-- February 2019 - USA Tables Added-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Glass Report</title>
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

' Collecting all of the information from the Inventory Tables and their Master Tables

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM QC_MASTER_GLASS ORDER BY ITEMNAME ASC"
'rs.Cursortype = GetDBCursorType
'rs.Locktype = GetDBLockType
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM QC_MASTER_SPACER ORDER BY ITEMNAME ASC"
Set rs2 = GetDisconnectedRS(strSQL2, DBConnection)

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM QC_MASTER_SEALANT ORDER BY ITEMNAME ASC"
Set rs3 = GetDisconnectedRS(strSQL3, DBConnection)

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * FROM QC_MASTER_MISC ORDER BY ITEMNAME ASC"
Set rs4 = GetDisconnectedRS(strSQL4, DBConnection)


if CountryLocation = "USA" then
	strSQL5 = "SELECT MasterId, Quantity, ConsumeDate FROM QC_GLASS_USA ORDER BY ID ASC"
	strSQL6 = "SELECT MasterId, ConsumeDate FROM QC_SPACER_USA ORDER BY ID ASC"
	strSQL7 = "SELECT MasterId, ConsumeDate FROM QC_SEALANT_USA ORDER BY ID ASC"
	strSQL8 = "SELECT MasterId, ConsumeDate FROM QC_MISC_USA ORDER BY ID ASC"

else
	strSQL5 = "SELECT MasterId, Quantity, ConsumeDate FROM QC_GLASS ORDER BY ID ASC"
	strSQL6 = "SELECT MasterId, ConsumeDate FROM QC_SPACER ORDER BY ID ASC"
	strSQL7 = "SELECT MasterId, ConsumeDate FROM QC_SEALANT ORDER BY ID ASC"
	strSQL8 = "SELECT MasterId, ConsumeDate FROM QC_MISC ORDER BY ID ASC"
end if

Set rs5 = Server.CreateObject("adodb.recordset")
Set rs5 = GetDisconnectedRS(strSQL5, DBConnection)

Set rs6 = Server.CreateObject("adodb.recordset")
Set rs6 = GetDisconnectedRS(strSQL6, DBConnection)

Set rs7 = Server.CreateObject("adodb.recordset")
Set rs7 = GetDisconnectedRS(strSQL7, DBConnection)

Set rs8 = Server.CreateObject("adodb.recordset")
Set rs8 = GetDisconnectedRS(strSQL8, DBConnection)

currentDate = Date
weekNumber = DatePart("ww", currentDate)
LastWeek = DatePart("ww", currentDate -7)
TwoAgoWeek = DatePart("ww", currentDate -14)


%> 
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
</head>
<body>
  <div class="toolbar">
        <h1 id="pageTitle">Glass Master Report</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_QC" target="_self">Glass<%response.write HomeSiteSuffix%></a>
    </div>  
   
   
        <ul id="Profiles" title="QC Report - Glass" selected="true">

<%
Dim Records
Dim ActiveTotal
'Dim TotalQuantity  Only 1 Pack per Active Record so this will always be the same as Records
Dim Consumed

		response.write "<li class='group'>QC GLASS REPORT </li>"
		response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		response.write "<li><table border='1' class='sortable' width='100%'><tr><th>Item Name</th><th>Manufacturer</th><th>Code</th><th>Lites/Pack</th><th>Size</th><th>Sqft(ft<sup>2</sup>)</th><th>$/Sqft</th><th>Entry Date</th><th>Actual Stock</th><th>Consumed</th><th> Extra Lites</th></tr>"
ActiveTotal = 0
do while not rs.eof
	if not rs5.bof then
		rs5.movefirst
	end if
	Records = 0
	'TotalQuantity = 0 
	Consumed = 0
	LastWeekConsumed = 0
	TwoAgoWeekConsumed = 0
	
	do while not rs5.eof
	
		if rs("ID") = rs5("MasterID") then
			if isnull(rs5("ConsumeDate")) then
				Records = Records + 1
				'TotalQuantity = TotalQuantity + rs5("Quantity)
			else
				Consumed = Consumed + 1
				if DatePart("ww", rs5("ConsumeDate")) = LastWeek then
					LastWeekConsumed = LastWeekConsumed + 1
				end if
				if DatePart("ww", rs5("ConsumeDate")) = TwoAgoWeek then
					TwoAgoWeekConsumed = TwoAgoWeekConsumed + 1
				end if
				
			end if
		end if
		
	rs5.movenext
	loop
	
		response.write "<tr><td>" & RS("ItemName") &"</td><td>" & RS("Manufacturer") & "</td><td>" & RS("Code") & "</td>"
		response.write "<td>" & RS("Pieces") &"</td><td>" & RS("Width") & " X " & RS("Height") & "</td><td>" & Round(RS("Width")*RS("Height")*RS("Pieces")/144,2) &"</td><td>$" & RS("Price") & "</td>"
		response.write "<td>" & RS("EntryDate") & "</td><td>" & Records & "</td><td>" & Consumed & " (" & LastWeekConsumed & " / " & TwoAgoWeekConsumed & ") </td><td>" & Rs("Lites") & "</td></tr>"
		ActiveTotal = ActiveTotal + Records
rs.movenext
loop
response.write "<li align = 'right'>Total Active Stock: " & ActiveTotal & "</li>"
ActiveTotal=0


response.write "</table></li>"

		response.write "<li class='group'>QC SPACER REPORT </li>"
		response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='30%'>Item Name</th><th  width='20%'>Manufacturer</th><th  width='20%'>Entry Date</th><th  width='15%'>Actual Stock</th><th  width='15%'>Consumed</th></tr>"


		
do while not rs2.eof
	if not rs6.bof then
		rs6.movefirst
	end if
	Records = 0 
	Consumed = 0
	LastWeekConsumed = 0
	TwoAgoWeekConsumed = 0
	
	do while not rs6.eof
	
		if rs2("ID") = rs6("MasterID") then
			if isnull(rs6("ConsumeDate")) then
				Records = Records + 1
			else
				Consumed = Consumed + 1
				if DatePart("ww", rs6("ConsumeDate")) = LastWeek then
					LastWeekConsumed = LastWeekConsumed + 1
				end if
				if DatePart("ww", rs6("ConsumeDate")) = TwoAgoWeek then
					TwoAgoWeekConsumed = TwoAgoWeekConsumed + 1
				end if
				
			end if
		end if
		
	rs6.movenext
	loop
	
		response.write "<tr><td>" & rs2("ItemName") &"</td><td>" & rs2("Manufacturer") & "</td><td>" & rs2("EntryDate") & "</td><td>" & Records & "</td><td>" & Consumed & " (" & LastWeekConsumed & " / " & TwoAgoWeekConsumed & ") </td></tr>"
ActiveTotal = ActiveTotal + Records
		rs2.movenext
loop
response.write "<li align = 'right'>Total Active Stock: " & ActiveTotal & "</li>"
ActiveTotal=0

response.write "</table></li>"

	response.write "<li class='group'>QC SEALANT REPORT </li>"
	response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='30%'>Item Name</th><th  width='20%'>Manufacturer</th><th  width='20%'>Entry Date</th><th  width='15%'>Actual Stock</th><th  width='15%'>Consumed</th></tr>"


		
do while not rs3.eof
	
	Records = 0 
	Consumed = 0
	LastWeekConsumed = 0
	TwoAgoWeekConsumed = 0

	if not rs7.bof then
		rs7.movefirst
	end if
	do while not rs7.eof
	
		if rs3("ID") = rs7("MasterID") then
			if isnull(rs7("ConsumeDate")) then
				Records = Records + 1
			else
				Consumed = Consumed + 1
				if DatePart("ww", rs7("ConsumeDate")) = LastWeek then
					LastWeekConsumed = LastWeekConsumed + 1
				end if
				if DatePart("ww", rs7("ConsumeDate")) = TwoAgoWeek then
					TwoAgoWeekConsumed = TwoAgoWeekConsumed + 1
				end if
				
			end if
		end if
		
	rs7.movenext
	loop
	
		response.write "<tr><td>" & rs3("ItemName") &"</td><td>" & rs3("Manufacturer") & "</td><td>" & rs3("EntryDate") & "</td><td>" & Records & "</td><td>" & Consumed & " (" & LastWeekConsumed & " / " & TwoAgoWeekConsumed & ") </td></tr>"
		ActiveTotal = ActiveTotal + Records
rs3.movenext
loop
response.write "<li align = 'right'>Total Active Stock: " & ActiveTotal & "</li>"
ActiveTotal=0

response.write "</table></li>"


	response.write "<li class='group'>QC MISCELLANEOUS REPORT </li>"
	response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='30%'>Item Name</th><th  width='20%'>Manufacturer</th><th  width='20%'>Entry Date</th><th  width='15%'>Actual Stock</th><th  width='15%'>Consumed</th></tr>"


		
do while not rs4.eof
	
	Records = 0 
	Consumed = 0
	if not rs8.bof then
		rs8.movefirst
	end if
	do while not rs8.eof
	
		if rs4("ID") = rs8("MasterID") then
			if isnull(rs8("ConsumeDate")) then
				Records = Records + 1
			else
				Consumed = Consumed + 1
			end if
		end if
		
	rs8.movenext
	loop
	
		response.write "<tr><td>" & rs4("ItemName") &"</td><td>" & rs4("Manufacturer") & "</td><td>" & rs4("EntryDate") & "</td><td>" & Records & "</td><td>" & Consumed & "</td></tr>"
		ActiveTotal = ActiveTotal + Records
rs4.movenext
loop
response.write "<li align = 'right'>Total Active Stock: " & ActiveTotal & "</li>"
ActiveTotal=0

response.write "</table></li>"

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3= nothing
rs4.close
set rs4 = nothing
rs5.close
set rs5 = nothing
rs6.close
set rs6 = nothing
rs7.close
set rs7 = nothing
rs8.close
set rs8 = nothing
DBConnection.close
set DBConnection = nothing

%>


</body>
</html>	