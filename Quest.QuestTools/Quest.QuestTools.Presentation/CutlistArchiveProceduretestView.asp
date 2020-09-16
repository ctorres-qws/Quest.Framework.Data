<!--#include file="dbpath_Quest_ArchiveLists.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			
			<!--CutlistArchiveProcedureMain.asp - is the Email code for sending this report-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest.mdb Archive Procedure</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>


<!--#include file="todayandyesterday.asp"-->
<% 
Server.ScriptTimeout=500
	currentDate = Date
	weekNumber = DatePart("ww", currentDate)
	OneWeekAgo = DateAdd("d",-7,currentDate)
	TwoWeekAgo = DateAdd("d",-14,currentDate)
	FourWeekAgo = DateAdd("d",-28,currentDate)
	CheckMinDate = DateAdd("yyyy",-5,currentDate)

 
'Collect TableNames from Schema Table 
Const adSchemaTables = 20
Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")

%>
</head>
<body>
<ul id="screen1" title="Quest Dashboard" selected="true">

	<li><b><u>CutList Archive Status: <%Response.write currentDate%> </u></b></li>
	<li><p>
	This email reflects archive update information from the last week as the first portion of information flow.<BR>
	Please respond to this email to Ariel Aziza and Michael Bernholtz with any cutlist issues this week.<BR>
	Specifically any cut-lists that could not be cut on the machines and had to be cut manually instead.<BR>
	</p></li>
	<li><b><i>Archive 1 - CUT</i></b></li>
	<table><TD>Table</TD><TD>Processed</TD><TD> Cut Number</TD><TD>Total NUmber</TD><TD> Cut Percentage</TD>
	
	
	
 <%



   

			TableNamePrefix = "STOP_*"
	rs.filter = "TABLE_NAME LIKE '" & TableNamePrefix & "' "
	 MostRecentDate = #01/01/1999#
	 DateCheck = #01/01/1999#
	Do while not rs.eof
	TableCount = TableCount + 1
		
		TableName = rs("TABLE_NAME")
		TableCheckStatus = FALSE
		
		Set Tablecheck = Server.CreateObject("adodb.recordset")
		if Left(TableName,3) = "ROD" or Left(TableName,4) = "STOP" then
			TC_SQL = "SELECT dDate, cStatus From [" & TableName & "]"
		else
			TC_SQL = "SELECT cDate, cStatus From [" & TableName & "]"
		end if
		Tablecheck.Cursortype = 1
		Tablecheck.Locktype = 3
		Tablecheck.Open TC_SQL, DBConnection
		
		StatusDone = 0
		StatusCount = 0
		MostRecentDate = #01/01/1999#
		
		'Record Count > 0 should exclude all Template pages
		'But will not catch tables processed with no data.
		if Tablecheck.RecordCount > 0 then
			Do while not TableCheck.eof
				StatusCount = StatusCount + 1
				if Tablecheck("cStatus") = True then
					StatusDone = StatusDone + 1
					'Most Recent Date calculated by checking each completed Date
					'Cutlist Archive Procedure Main - has 4 values to compare
					'currentDate = Date
					'OneWeekAgo = DateAdd("ww",-1,currentDate)
					'TwoWeekAgo = DateAdd("ww",-2,currentDate)
					'FourWeekAgo = DateAdd("ww",-4,currentDate)
					
					if Left(TableName,3) = "ROD" or Left(TableName,4) = "STOP" then
						
						if len(Tablecheck("dDate"))>5 then
							if InStr(Tablecheck("dDate"),".") = 5 then
								Dateyear = Left(Tablecheck("dDate"),4)
								Datemonth = Mid(Tablecheck("dDate"),6,2)
								DateDay = Right(Tablecheck("dDate"),2)
								DateCheck2 = DateMonth & "/" & DateDay& "/" & DateYear
							end if 
							if InStr(Tablecheck("dDate"),".") = 3 then
								Dateyear = Right(Tablecheck("dDate"),4)
								Datemonth = Mid(Tablecheck("dDate"),4,2)
								DateDay = Left(Tablecheck("dDate"),2)
								DateCheck2 = DateMonth & "/" & DateDay& "/" & DateYear
							end if 						
						
						
							if isdate(DateCheck2) = True then
								DateCheck = DateCheck2
							end if
						end if
					else
				Response.write Tablecheck("cDate") & IsDate(Tablecheck("cDate"))
						if Isdate(Tablecheck("cDate")) Then
							DateCheck = CDATE(Tablecheck("cDate"))
							
						end if
					end if
					
					if DateCheck > MostRecentDate Then
						MostRecentDate = DateCheck
					end if
					
				end if
			Tablecheck.movenext
			Loop
			
			Response.write "<TR>"
			Response.write "<TD>" & rs("TABLE_NAME") & "</TD>"
			Response.write "<TD>" & MostRecentDate & "</TD>"
			Response.write "<TD>" & StatusDone & " </TD>"
			Response.write "<TD>" & StatusCount & " </TD>"
			Response.write "<TD>" & StatusDone / StatusCount * 100 & " </TD>"
			Response.write "</TR>"
		
		end if
		Tablecheck.close
		Set Tablecheck = nothing
	
		
	rs.movenext	
	Loop
%>
</table></ul>
  
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
