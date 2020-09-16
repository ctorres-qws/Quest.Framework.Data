<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="connect_tablist.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>

    </head>
<body>

    <ul id="Profiles" title="Kittings Processed" selected="true">
<% 
response.write "<li class='group'>Shift Frame Kits Processed but not yet Kitted</li>"
response.write "<li><table border='1' class='sortable'><tr><th>JOB</th><th>Floor</th><th id ='ProDay'>Processed Date</th><th>Kitting Status</th><th>Days old</th></tr>"

rs.filter = ""
rs.filter = "TABLE_NAME LIKE '%SHIFT_%'"

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Z_Cutlists WHERE Cutlist Like 'SHIFT%' AND YEAR >= 2020 ORDER BY ID DESC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT * FROM X_SHIFTHARDWARE ORDER BY ID DESC"
	rs3.Cursortype = GetDBCursorType
	rs3.Locktype = GetDBLockType
	rs3.Open strSQL3, DBConnection

TodayYear = DateAdd("yyyy", -2, Date)
TodayMonth = DateAdd("m", -2, Date)
TodayDay= DateAdd("d", -2, Date)
TodayDate = TodayYear & TodayMonth &TodayDay


do while not rs.eof

	TableName = RIGHT(rs("Table_Name"), Len(rs("Table_Name"))-6)
	JobName = Left(TableName,3)
	FloorName = Right(TableName, Len(TableName)-3)

	rs2.Filter = ""
	rs2.Filter = "JOB ='" & JobName & "' AND FLOOR = '" & FloorName & "'"
	if rs2.eof then
		YearNum = 1000
		ProcessDate = "N/A"
	else
		YearNum = rs2("Year") + 0
		MonthNum =rs2("Month") + 0
		if Len(MonthNum) = 1 then
			MonthNum = "0" & MonthNum
		end if
		DayNum = rs2("Day") + 0
		if Len(DayNum) = 1 then
			DayNum = "0" & DayNum
		end if
		
		ProcessDate = FormatDateTime(YearNum & "/" & MonthNum & "/" & DayNum,2)
	end if

	rs3.filter = ""
	rs3.filter = "JOB = '" & JobName & "' AND FLOOR = '" & FloorName & "'"
	DisplayKit = FALSE
	KittingDate = ""
	if rs3.eof then
		KittingDate = "No Kitting Data Available"
		DisplayKit = TRUE
	else
		KittingDate = Rs3("CompletedDate")
		if (Len(KittingDate) < 2) or (isnull(KittingDate)=True) then
			DisplayKit = TRUE
			KittingDate = "Not Yet Kitted"
		end if
	end if
	if DisplayKit = True then
	'Clean up
		if Jobname = "AAA" or JobName = "ARI" or JobName = "DAN" or JobName = "MAX" or JobName = "_Te" or (YearNum < 2020 and YearNum <> 1000) then
		else
				ProcessDay = YearNum &  MonthNum & DayNum
				Response.write "<TR>"
				Response.write "<TD>" & JobName & "</TD>"
				Response.write "<TD>" & FloorName & "</TD>"
				Response.write "<TD sorttable_customkey='" & ProcessDay & "' >" & ProcessDate & "</TD>"
				Response.write "<TD>" & KittingDate & "</TD>"
				if ProcessDate > TodayDate then
					Response.write "<TD></TD>"
				else
					Response.write "<TD>Within 10 Days</TD>"
				end if
			
			end if
	end if
	
	Response.write "</TR>"

rs.movenext
loop

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close 
Set rs3= nothing
DBConnection.close 
set DBConnection = nothing

%>
      </ul>

</body>
</html>
