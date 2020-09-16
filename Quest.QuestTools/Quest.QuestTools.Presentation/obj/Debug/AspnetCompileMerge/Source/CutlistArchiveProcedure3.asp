<!-- Stored Procedure include to be run as part of CutlistArchiveProcedureMain-->
<!--CutListArchiveProcedure1 - CUT/HCUT/DMSAW/ROD -->
<!--CutListArchiveProcedure2 - DMSDR/SHIFT/R3/STOP -->
<!--CutListArchiveProcedure3 - SCRN-->
<!--CutListArchiveProcedure4 - QSU/QSP/PANEL-->
<!--Michael Bernholtz - September 2019: Approved by Ariel Aziza -->

 <%

 
'Collect TableNames from Schema Table 
'Const adSchemaTables = 20
'Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")

 TableCount = 0
 Table100 = 0
 TableSix = 0
 TableDeleted = 0



i = 1
   
Do Until i >= 2
 TableCount = 0
 Table100 = 0
 TableSix = 0
 TableDeleted = 0
 RecentTables = ""
   
	Select Case i
		Case 1
			TableNamePrefix = "SCRN_*"
	End Select

   
	rs.filter = "TABLE_NAME LIKE '" & TableNamePrefix & "' "
	 
	Do while not rs.eof
	TableCount = TableCount + 1
		
		TableName1 = rs("TABLE_NAME")
		TableName2 = "SCRNPROD_" & Right(rs("TABLE_NAME"),Len(rs("TABLE_NAME")) -5)
		
		TableCheckStatus = FALSE
		
		Set Tablecheck = Server.CreateObject("adodb.recordset")
			TC_SQL = "SELECT dDate, cStatus From [" & TableName1 & "]"
		Tablecheck.Cursortype = 1
		Tablecheck.Locktype = 3
		Tablecheck.Open TC_SQL, DBConnection
		
		StatusDone = 0
		StatusCount = 0
		MostRecentDate = #01/01/1999#
		currentDate = Date
		SixWeekAgo = DateAdd("ww",-6,currentDate)
		
		'Record Count > 0 should exclude all Template pages
		'But will not catch tables processed with no data.
		if Tablecheck.RecordCount > 0 then
			Do while not TableCheck.eof
				StatusCount = StatusCount + 1
				if Tablecheck("cStatus") = True then
					StatusDone = StatusDone + 1
					if Isdate(Tablecheck("dDate")) Then
						DateCheck = CDATE(Tablecheck("dDate"))
						if DateCheck > MostRecentDate Then
							MostRecentDate = DateCheck
						end if
					end if
					
				end if
			Tablecheck.movenext
			Loop
			
			if StatusDone = StatusCount then
			'100%
				TableCheckstatus = TRUE
				Table100 = Table100 + 1
			end if
			
			if (MostRecentDate > CheckMinDate) AND (MostRecentDate < SixWeekAgo) AND statusDone > 0  then
			'Older than 6 weeks but started
				TableCheckstatus = TRUE
				TableSix = TableSix + 1
			end if
		
		end if
		Tablecheck.close
		Set Tablecheck = nothing
		

		if TableCheckstatus = FALSE then
		else

			Set rs2 = Server.CreateObject("adodb.recordset")
			strSQL2 = "Select * into [MS Access;DATABASE=f:\database\ArchiveLists.mdb]." & TableName1 &  " FROM [MS Access;DATABASE=f:\database\quest.mdb]." & TableName1
				On Error Resume Next  
			rs2.Open strSQL2, DBConnection2
				On Error GoTo 0
				
		
			SQL3 = "Drop TABLE " & TableName1 
				On Error Resume Next  
			set RS3 = DBConnection.Execute(SQL3)
			if Err.Number = 0 then
				TableDeleted = TableDeleted + 1
			end if 

				On Error GoTo 0
			
			SQL4 = "UPDATE Z_Cutlists SET ACTIVE = FALSE WHERE CUTLIST = '" & TableName1 &  "'"
				On Error Resume Next  
			RS4 = DBConnection.Execute(SQL4)
				On Error GoTo 0
				
						Set rs2 = Server.CreateObject("adodb.recordset")
			strSQL2 = "Select * into [MS Access;DATABASE=f:\database\ArchiveLists.mdb]." & TableName2 &  " FROM [MS Access;DATABASE=f:\database\quest.mdb]." & TableName2
				On Error Resume Next  
			rs2.Open strSQL2, DBConnection2
				On Error GoTo 0
				
		
			SQL3 = "Drop TABLE " & TableName2
				On Error Resume Next  
			set RS3 = DBConnection.Execute(SQL3)
			if Err.Number = 0 then
				TableDeleted = TableDeleted + 1
			end if 

				On Error GoTo 0
			
			SQL4 = "UPDATE Z_Cutlists SET ACTIVE = FALSE, ArchiveDate = '" & currentDate & "' WHERE CUTLIST = '" & TableName2 &  "'"
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
	<li>Tables Started but old than 6 Weeks: <%response.write RecentTables %></li>
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
