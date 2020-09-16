<!-- Stored Procedure include to be run as part of CutlistArchiveProcedureMain-->
<!--CutListArchiveProcedure1 - CUT/HCUT/DMSAW/ROD -->
<!--CutListArchiveProcedure2 - DMSDR/SHIFT/R3/STOP -->
<!--CutListArchiveProcedure3 - SCRN-->
<!--CutListArchiveProcedure4 - QSU/QSP/PANEL-->
<!--Michael Bernholtz - September 2019: Approved by Ariel Aziza -->
<!--Date: February 10, 2020
	Modified By: Michelle Dungo
	Changes: Modified to remove extra reserved word 'else'
-->
 <%

 
'Collect TableNames from Schema Table 
'Const adSchemaTables = 20
'Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")

 TableCount = 0
 Table100 = 0
 Table90 = 0
 Table50 = 0
 TableDeleted = 0



i = 1
   
Do Until i >= 5
 TableCount = 0
 Table100 = 0
 Table90 = 0
 Table50 = 0
 TableDeleted = 0
 RecentTables = ""
   
	Select Case i
		Case 1
			TableNamePrefix = "DMSDR_*"
		Case 2
			TableNamePrefix = "SHIFT_*"
		Case 3
			TableNamePrefix = "STOP_*"
		Case 4
			TableNamePrefix = "R3PANEL_*"
	End Select

   
	rs.filter = "TABLE_NAME LIKE '" & TableNamePrefix & "' "
	 
	Do while not rs.eof
	TableCount = TableCount + 1
		
		TableName = rs("TABLE_NAME")
		TableCheckStatus = FALSE
		
		Set Tablecheck = Server.CreateObject("adodb.recordset")
		if Left(TableName,4) = "STOP" then
			TC_SQL = "SELECT dDate, cStatus From [" & TableName & "]"
		else
			TC_SQL = "SELECT cDate, cStatus From [" & TableName & "]"
		end if
		Tablecheck.Cursortype = 1
		Tablecheck.Locktype = 3
		Tablecheck.Open TC_SQL, DBConnection
		
		StatusDone = 0
		StatusCount = 0
		DateCheck = #01/01/2000#
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
					
					if Left(TableName,4) = "STOP" then
					'	if Isdate(Tablecheck("dDate")) Then
					'		DateCheck = CDATE(Tablecheck("dDate"))
					'	end if
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
						if Isdate(Tablecheck("cDate")) Then
							DateCheck = CDATE(Tablecheck("cDate"))
						end if
					end if
					
					if CDATE(DateCheck) > CDATE(MostRecentDate) Then
						MostRecentDate = DateCheck
					end if
					
				end if
			Tablecheck.movenext
			Loop
			
			if StatusDone = StatusCount then
			'100%
				TableCheckstatus = TRUE
				Table100 = Table100 + 1
			end if
			


				if ((StatusDone/StatusCount) >= 0.9) AND (StatusDone >= 1)  AND (StatusDone <> StatusCount) then
				'90%
					
					if (TwoWeekAgo > MostRecentDate) AND (MostRecentDate > CheckMinDate) Then
						TableCheckstatus = TRUE
						Table90 = Table90 + 1
					end if
				end if
				
				if ((StatusDone/StatusCount) < 0.5) AND (StatusDone >= 1) AND (CDATE(MostRecentDate) > CDATE(CheckMinDate)) then
				'0-50%%
					
					if FourWeekAgo > MostRecentDate Then
						TableCheckstatus = TRUE
						Table50 = Table50 + 1
					else
						RecentTables = RecentTables & " " & TABLE_NAME
					end if
				end if

		
		end if
		Tablecheck.close
		Set Tablecheck = nothing
		

		if TableCheckstatus = FALSE then
		else

			Set rs2 = Server.CreateObject("adodb.recordset")
			strSQL2 = "Select * into [MS Access;DATABASE=f:\database\ArchiveLists.mdb]." & TableName &  " FROM [MS Access;DATABASE=f:\database\quest.mdb]." & TableName
				On Error Resume Next  
			rs2.Open strSQL2, DBConnection2
				On Error GoTo 0
				
		
			SQL3 = "Drop TABLE " & TableName 
				On Error Resume Next  
			set RS3 = DBConnection.Execute(SQL3)
			if Err.Number = 0 then
				TableDeleted = TableDeleted + 1
			end if 

				On Error GoTo 0
			
			SQL4 = "UPDATE Z_Cutlists SET ACTIVE = FALSE, ArchiveDate = '" & currentDate & "' WHERE CUTLIST = '" &TableName &  "'"
				On Error Resume Next  
			RS4 = DBConnection.Execute(SQL4)
				On Error GoTo 0

		end if
		
	rs.movenext	
	Loop

%>
   
	<li><B><U><%Response.write TableNamePrefix %></U></B></li>
	<li>Tables Counted: <%response.write TableCount %></li>
	<li>Tables at 100: <%response.write Table100 %></li>
	<li>Tables at 90 (Older than 2 Weeks): <%response.write Table90 %></li>
	<li>Tables at 0-50 (Older than 4 Weeks): <%response.write Table50 %></li>
	<li>Tables at 0-50 (Newer than 4 Weeks): <%response.write RecentTables %></li>
	<li>Tables Archived: <%response.write TableDeleted %></li>
	<li>Tables Remaining: <%response.write TableCount - TableDeleted %></li>

<%

i = i +1
loop

'	rs.close
'	set rs= nothing

'DBConnection.close
'set DBConnection = nothing
'DBConnection2.close
'set DBConnection2 = nothing
 %>

<!--
</body>
</html>
-->
