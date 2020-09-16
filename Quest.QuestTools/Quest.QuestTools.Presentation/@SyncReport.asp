<!--#include file="@common.asp"-->
<html>
	<head>
		<style>
			body { font-family: arial; font-size: 11px;}
			div { font-size: 12px; font-weight: bold;}
			.csWarning { background-color: yellow; }
		</style>
	</head>
	<body>
<%
	gbMyDebug = False
	Response.Write("SQL: " & GetConnectionStr(true) & "<br/>")
	Response.Write("Access: " & GetConnectionStr(false) & "<br/><br/>")

	Dim gdt_Now: gdt_Now = Now

	Dim a_Tables
	str_SyncURL = "@ReSyncTable.asp?Table={0}&Action=RESYNC&Submit=ReSync"
	str_SyncURL_ID = "@ReSyncTable.asp?Table={0}&Identity=on&Action=RESYNC&Submit=ReSync"
	Header("Processing System Tables")
	SectionStart
	a_Tables = Array("LabelExport","GlassTypes","SPBase_PickList","Styles","Styles_DEF","Styles_Active","StylesPanel","X_Win_Prod","XQSP_GlassTypes","XQSU_GlassTypes","XQSU_OTSpacer","Y_Entry","Y_Color","Z_Jobs","Z_Floors","Y_Entry")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("Processing Inventory Tables")
	SectionStart
	a_Tables = Array("Y_Inv","Y_InvLog","Y_Master","Y_Hardware","Y_Hardware_Log","Y_Hardware_Master","ZShift_Template","X_ShiftHardware")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("Processing Glass Tables")
	SectionStart
	a_Tables = Array("Z_GlassDB")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("Processing Jobs Tables")
	SectionStart
	a_Tables = Array("Z_Jobs")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("Report Tables")
	a_Tables = Array("V_Report1","V_Report2","V_Report3")
	SectionStart
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("System Tables")
	a_Tables = Array("X_Employees")
	SectionStart
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	Header("System Scanning Tables")
	a_Tables = Array("X_Barcode","X_BarcodeGA","X_BarcodeOV","X_BarcodeP","X_Glazing","X_Shipping","X_Ship")
	SectionStart
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next

	a_Tables = Array("X_Shipping_Truck")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	
	a_Tables = Array("X_Ship_Truck")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next

	SectionEnd

	Header("Processing System Tables")
	a_Tables = Array("X_Barcode","X_BarcodeGA","X_BarcodeOV","X_BarcodeP","X_Glazing","X_Shipping","X_Ship")
	SectionStart
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, true)
	Next

	a_Tables = Array("X_Shipping_Truck")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	
	a_Tables = Array("X_Ship_Truck")
	For i = 0 To UBound(a_Tables)
		Call LogTable(a_Tables(i),0, false)
	Next
	SectionEnd

	'Header("Report Tables (Ignore)")
	'a_Tables = Array("ProDmt1","ProDmt2","ProEcoHor","ProEcoHorWork","ProEcoVert","ProEcoVert2","ProQHor","ProstopStatus","ProStopStatus2","ProStopStatus3","ProStopStatus4","ProZipperBlue","ProZipperRed")
	'SectionStart
	'For i = 0 To UBound(a_Tables)
	'	Call LogTable(a_Tables(i),0, false)
	'Next
	'SectionEnd

	Sub LogTable(str_Table,i_Day, b_DateTime)
		Dim cn_SQL, rs_SQL
		Dim cn_Access, rs_Access
		Dim dt_Now
		Dim str_SQL
		dt_Now = DateAdd("d",i_Day,gdt_Now)

		str_DateField = "DateTime"

		str_SQL = "SELECT COUNT(*) as Counts FROM [" & str_Table & "] "
		If b_DateTime Then
			str_SQL = str_SQL & " WHERE Year({0})=" & Year(dt_Now) & " AND Month({0})=" & Month(dt_Now) & " AND Day({0}) = " & Day(dt_Now)
		End If

		If UCase(str_Table) = "X_SHIPPING" Or UCase(str_Table) = "X_SHIPPING_TRUCK" Or UCase(str_Table) = "X_SHIP" Or UCase(str_Table) = "X_SHIP_TRUCK" Then
			str_DateField = "ShipDate"
		End If

		If UCase(str_Table) = "Y_INVLOG" Then
			str_SQL = str_SQL & " WHERE YEAR > 2016 "
		End If

		str_SQL = Replace(str_SQL, "{0}", str_DateField)

		Set cn_SQL = Server.CreateObject("ADODB.Connection")
		cn_SQL.ConnectionString = GetConnectionStr(true)
		cn_SQL.Open

		Set cn_Access = Server.CreateObject("ADODB.Connection")
		If UCase(str_Table) = "USERS" Then
			cn_Access.ConnectionString = GetConnectionStrAdmin(false)
		Else
			cn_Access.ConnectionString = GetConnectionStr(false)
		End If
		cn_Access.Open

		Set rs_Access = Server.CreateObject("ADODB.Recordset")

		rs_Access.Cursortype = GetDBCursorType
		rs_Access.Locktype = GetDBLockType
		rs_Access.Open str_SQL, cn_Access

		Set rs_SQL = Server.CreateObject("ADODB.Recordset")
		rs_SQL.Cursortype = GetDBCursorType
		rs_SQL.Locktype = GetDBLockType
		rs_SQL.Open str_SQL, cn_SQL
		str_Class = ""
		If rs_Access("Counts") <> rs_SQL("Counts") Then
			str_Class = "csWarning"
		End If

		Select Case(UCase(str_Table))
			Case "GLASSTYPES","STYLES","STYLES_DEF","STYLESPANEL","V_REPORT1","V_REPORT2","V_REPORT3","Z_JOBS","Z_FLOORS", "Y_MASTER","XQSU_GLASSTYPES","XQSP_GLASSTYPES","XQSU_OTSPACER","X_WIN_PROD","X_EMPLOYEES","X_SHIPPING_TRUCK","Y_HARDWARE","Y_HARDWARE_LOG","Y_HARDWARE_MASTER","Y_ENTRY","Y_COLOR","PRODMT1","PRODMT2","PROECOHOR","PROECOVERT","PROECOHORWORK","PROECOVERT2","PROQHOR","PROSTOPSTATUS","PROSTOPSTATUS2","PROSTOPSTATUS3","PROSTOPSTATUS4","PROZIPPERBLUE","PROZIPPERRED","X_SHIP_TRUCK","X_SHIP"
				str_ReSync = "<a target='SyncTable' href='" & Replace(str_SyncURL,"{0}", str_Table) & "'>ReSync</a>"
			Case "LABELEXPORT","STYLES_ACTIVE","X_BARCODE_LINEITEM"
				str_ReSync = "<a target='SyncTable' href='" & Replace(str_SyncURL_ID,"{0}", str_Table) & "'>ReSync</a>"
		End Select

		Response.Write("<tr class='" & str_Class & "'><td>" & str_Table & "</td><td style='text-align: right; padding-right: 5px;'>" & rs_Access("Counts") & "</td><td style='text-align: right; padding-right: 5px;'>" & rs_SQL("Counts") & " </td><td> " & str_ReSync & "</td></tr>")

		rs_SQL.Close
		Set rs_SQL = Nothing

		rs_Access.Close
		Set rs_Access = Nothing

		cn_SQL.Close
		cn_Access.Close

	End Sub

	Sub SectionStart()
		Response.Write("<table style='border:1px solid #CCCCCC;'>")
		Response.Write("<tr style='background-color: #eaeaea'><td>Table</td><td>Counts(Access)</td><td>Counts(SQL)</td><td></td></tr>")
	End Sub

	Sub SectionEnd()
		Response.Write("</table>")
	End Sub

	Sub Header(str_Header)
		Response.Write("<br /><div>" & str_Header & "</div>")
	End Sub

	Function GetConnectionStrAdmin(b_SQL)
		GetConnectionStrAdmin = gstr_DB_Access_Admin
	End Function

%>
</body>
</html>