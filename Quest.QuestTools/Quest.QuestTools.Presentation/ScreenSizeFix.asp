<HTML>
<BODY>
<!--Screen Size fix - November 2019-->
<%
'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN = DSN & "DBQ=F:\database\quest.mdb"
DSN = DSN & ";PWD=stewart"
DBConnection.Open DSN

'Collect TableNames from Schema Table 
Const adSchemaTables = 20
Set rs = DBConnection.OpenSchema(adSchemaTables)
'rs("TABLE_NAME")
%>
 <%


	TableNamePrefix = "SCRN_*"
   
	rs.filter = "TABLE_NAME LIKE '" & TableNamePrefix & "' "
	 
	Do while not rs.eof
	TableCount = TableCount + 1
		
		TableName1 = rs("TABLE_NAME")
		TableName2 = "SCRNPROD_" & Right(rs("TABLE_NAME"),Len(rs("TABLE_NAME")) -5)
		
		TableCheckStatus = FALSE
		
		Set Tablecheck = Server.CreateObject("adodb.recordset")
			TC_SQL = "SELECT * From [" & TableName1 & "]"
		Tablecheck.Cursortype = 1
		Tablecheck.Locktype = 3
		Tablecheck.Open TC_SQL, DBConnection
		
		Set Tablecheck2 = Server.CreateObject("adodb.recordset")
			TC_SQL2 = "SELECT * From [" & TableName2 & "]"
		Tablecheck2.Cursortype = 1
		Tablecheck2.Locktype = 3
		Tablecheck2.Open TC_SQL2, DBConnection
		
		
		'Record Count > 0 should exclude all Template pages
		'But will not catch tables processed with no data.
		if Tablecheck.RecordCount > 0 then
			Do while not TableCheck.eof
			
			
			JOB = Tablecheck("JOB")
			FLOOR = Tablecheck("FLOOR")
			TAG = Tablecheck("TAG")
			SIDE =  Trim(Tablecheck("AWNCAS"))
			SideCheck = ""
			if Side = "W" then
				SideCheck = "CW"
			else 
				SideCheck = "CH"
			end if
			
			Tablecheck2.filter = " JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG = '" & TAG & "'"
			
			
			
			Response.write "<LI>"
			Response.write JOB & FLOOR & TAG & SIDE & "::" & TableCheck("Length") & "::"
			Response.write TableCheck2("JOB") & TableCheck2("FLOOR") & TableCheck2("TAG") & TableCheck2(SideCheck)
			
			WrongValue = (TableCheck2(SideCheck) + 0) - (TableCheck("Length") + 0)
			if WrongValue = 0 then 
			else
			Response.write " -- - - - - - - - - - " & WrongValue
			end if
			Response.write "</LI>"
		
		'	if WrongValue <> 0 then
		'		TableCheck("Length") = TableCheck2(SideCheck)
		'		Tablecheck.update
		'	end if
			
		
			Tablecheck.movenext
			Loop
			
			
		
		end if
		Tablecheck.close
		Set Tablecheck = nothing
		
		Tablecheck2.close
		Set Tablecheck2 = nothing
	
		
	rs.movenext	
	Loop

%>
   

<%



	rs.close
	set rs= nothing

DBConnection.close
set DBConnection = nothing

 %>

</body>
</html>

