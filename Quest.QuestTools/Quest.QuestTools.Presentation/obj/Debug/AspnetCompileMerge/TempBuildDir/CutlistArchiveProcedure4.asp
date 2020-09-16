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
 TableFull = 0
 TableDeleted = 0

i = 1
   
Do Until i >= 4
   
	Select Case i
		Case 1
			TableNamePrefix = "QSU_*"
		Case 2
			TableNamePrefix = "QSP_*"
		Case 3
			TableNamePrefix = "PANEL_*"

	End Select

   
	rs.filter = "TABLE_NAME LIKE '" & TableNamePrefix & "' "
 
	Do while not rs.eof

		TableCount = TableCount + 1
		TableName = rs("TABLE_NAME")

		if UCASE(TableName) = "PANEL_TEMPLATE" then
		else
			
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
	<li>Tables Deleted: <%response.write TableDeleted %></li>
	<li>Tables Remaining: <%response.write TableCount - TableDeleted %></li>

	
 <%
 
i = i+ 1
Loop

'rs.close
'set rs= nothing

'DBConnection.close
'set DBConnection = nothing
'DBConnection2.close
'set DBConnection2 = nothing
 
%>
<!--
</body>
</html>
-->