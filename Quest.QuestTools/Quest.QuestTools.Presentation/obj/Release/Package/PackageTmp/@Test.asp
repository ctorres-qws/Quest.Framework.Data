<!--#include file="@common.asp"-->
<%

'	Reporttime = "previous"
'	currentDate = Date()

'	Call SendEmail("hsandhu@questwindows.com", "Test Subject", "Test Message")

' Sets the Month and Year of the Data to be Saved from Y_INV
'	Select Case Reporttime
'		Case "current"
'			SnapMonth = Month(now)
'			SnapYear = Year(now)
'		Case "previous"
'			dt_Now = Now
'			str_SnapDate = "1-" & ga_Months(Month(dt_Now)) & "-" & Year(dt_Now)
'			If Month(now) = 1 Then
'				SnapMonth =  12
'				SnapYear = Year(now) -1
'			Else
'				SnapMonth = Month(now)-1
'				SnapYear = Year(now)
'			End If
'		Case Else
'			SnapMonth = Month(now)
'			SnapYear = Year(now)
'	End Select
'
'' Adds a 0 to Num 1-9 for consistency 
'	If SnapMonth < 10 Then
'		SnapMonth = "0" & SnapMonth
'	End If
'
'	Set rs2 = Server.CreateObject("adodb.recordset")
'	strSQL2 = "Select * into Y_INV" & SnapMonth & SnapYear & " FROM " & gstr_SQLDB_Primary & ".[Y_INV] WHERE DateIN < #" & str_SnapDate & "#"
'	Response.Write(strSQL2 & "<br/>")
'	'rs2.Open strSQL2, DBConnection2
'	'set rs2 = nothing
'
'	Set rs5 = Server.CreateObject("adodb.recordset")
'	strSQL5 = "Select * into Y_Hardware" & SnapMonth & SnapYear & " FROM " & gstr_SQLDB_Primary & ".[Y_Hardware] WHERE EnterDate < #" & str_SnapDate & "#"
'	Response.Write(strSQL5 & "<br/>")
'	'rs5.Open strSQL5, DBConnection2
'	'set rs5 = nothing
'
'	Set rs3 = Server.CreateObject("adodb.recordset")
'	strSQL3 = "Select * into X_Barcode" & SnapMonth & SnapYear & " FROM " & gstr_SQLDB_Primary & ".[X_BARCODE] WHERE MONTH = " & SnapMonth & "AND YEAR = " & SnapYear 
'	'rs3.Open strSQL3, DBConnection2
'	'set rs3 = nothing

'***************************************** PREF CODE

	Response.Write("<br />")

	If Request("T") = 1 Then
		'ReRun(36772)
		Response.Write("Add Inventory")
	ElseIf Request("R") = 1 Then
		Response.Write("Reverse Inventory")
		ReverseInventory("16941") '17352,17354,17398,17337,17338,17324,16941
	ElseIf Request("RS") = 1 Then
		Response.Write("Warehouse In:" & Request("Warehouse_In") & "<br/>")
		Response.Write("Warehouse Out: " & Request("Warehouse_Out") & "<br/>")
		Response.Write("Part: " & Request("Part") & "<br/>")
		Response.Write("Qty: " & Request("Qty") & "<br/>")
		Response.Write("Colour: " & Request("Colour") & "<br/>")
		Response.Write("Length: " & Request("PartLength") & "<br/>")

		Response.Write("<br/><br/>")

		Call PrefAddTransfer(0, UCase(Request("Warehouse_Out")), UCase(Request("Warehouse_In")), Request("Part") & "", Request("Qty") & "", Request("Colour"), CLng(Request("PartLength")), true, "Adjustment", "", "", "")

		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-146", 1, "UC114626", 18, true, "Adjustment", "", "", "")

	ElseIf Request("M") = 1 Then
		' 6705.6 - 22
		' 6400.8 - 21
		' 6096 - 20
		' 5486.4 - 18

		Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-146", 1, "UC114626", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "NC70000", 180, "LIL-Ext", 21.3, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-164", 150, "VNY-Ext", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-164", 20, "MILL", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-164", 20, "MILL", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-164", 20, "MILL", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-162", 70, "VNY-Ext", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-151", 6, "MILL", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "DURAPAINT", "ADJUSTMENT", "Que-151", 6, "UC70123F", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-151", 6, "UC70123F", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-145", 1, "UC82989XL", 18, true, "Adjustment", "", "", "") 'SUB
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-145", 11, "UC70123F", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-141", 10, "IKE-Ext", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-145", 369, "K1285", 18, true, "Adjustment", "", "", "") 'SUB
		'Call PrefAddTransfer(0, "ADJUSTMENT", "HORNER", "Que-145", 369, "MILL", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-144", 180, "K1285", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-145", 224, "MILL", 18, true, "Adjustment", "", "", "") 'SUB
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-144", 180, "K1285", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-17", 11, "MILL", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-143", 71, "UC82989XL", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-17", 2, "OTX-Ext", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "A-201", 220, "VNY-INT", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "A-202", 10, "VNY-INT", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-189", 210, "VNY-EXT", 20, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-203", 221, "CLR-EXT", 20, true, "Adjustment", "", "", "") 'SUBTRACT - DID NOT WORK - CAN'T FIND COLOUR
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-28", 1, "SJU-INT", 18, true, "Adjustment", "", "", "") 'SUBTRACT
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-60", 3, "UC119705", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-62", 1, "SJU-INT", 18, true, "Adjustment", "", "", "") 'ADD
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-62", 1, "SJU-INT", 18, true, "Adjustment", "", "", "") 'ADD TO DURAPAINT
		'Call PrefAddTransfer(0, "HORNER", "ADJUSTMENT", "QUE-94", 827, "MILL", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-43", 16, "MILL", 20, true, "Adjustment", "", "", "") 'ADD TO DURAPAINT
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-61", 2, "GBK-Ext", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA

		'Call PrefAddTransfer(17555, "NASHUA", "ADJUSTMENT", "QUE-3", 180, "K1285", 6, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(17555, "ADJUSTMENT", "NASHUA", "QUE-108", 82, "UC114626", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(17555, "NASHUA", "ADJUSTMENT", "QUE-108", 82, "UC114626", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(31870, "NASHUA", "ADJUSTMENT", "QUE-84", 80, "UC70192F", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(31334, "ADJUSTMENT", "NASHUA", "QUE-172", 4, "GBK-EXT", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "NC70049", 5, "WLL-EXT", 21.33, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "QUE-164", 7, "UC113883", 21.33, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "NC70006", 12, "MILL", 21.33, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-142", 2, "UC114626", 20, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-143", 390, "Mill", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-163", 326, "ANY", 22, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-163", 326, "ANY", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-163", 280, "MILL", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-166", 234, "MILL", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-28", 2, "SJX-Ext", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-62", 413, "SJU-Int", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "HORNER", "ADJUSTMENT", "QUE-166", 234, "ANY", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "HORNER", "ADJUSTMENT", "QUE-163", 280, "ANY", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "HORNER", "ADJUSTMENT", "QUE-163", 280, "MILL", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-18-148-S", 127, "Mill", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "HORNER", "Que-166", 467, "ANY", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "HORNER", "ADJUSTMENT", "QUE-163", 280, "ANY", 22, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-147", 3, "UC114626", 20, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-144", 9, "PAM-Int", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "QUE-144", 1, "SJX-Int", 18, true, "Adjustment", "", "", "")
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "NC70001", 88, "Mill", 21.33, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-143", 70, "UC82989XL", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-17", 18, "Mill", 20, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-17", 18, "Mill", 20, true, "Adjustment", "", "", "") 
		'Call PrefAddTransfer(0, "ADJUSTMENT", "DURAPAINT", "Que-17", 18, "Mill", 20, true, "Adjustment", "", "", "") 
		'Call PrefAddTransfer(0, "DURAPAINT", "ADJUSTMENT", "Que-17", 18, "Mill", 20, true, "Adjustment", "", "", "") 
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-202", 120, "CLEAR AND-EXT", 20, true, "Adjustment", "", "", "") 
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-202", 120, "CLEAR AND-EXT", 20, true, "Adjustment", "", "", "") 
		'Call PrefAddTransfer(0, "ADJUSTMENT", "NASHUA", "Que-94", 20, "OTX-Ext", 18, true, "Adjustment", "", "", "") 'ADD TO NASHUA
		'Call PrefAddTransfer(0, "NASHUA", "ADJUSTMENT", "Que-64", 9, "BNA-Ext", 18, true, "Adjustment", "", "", "")
		
		
	End If

	Function ReverseInventory(str_ID)
		Dim cn_DB: Set cn_DB = Server.CreateObject("ADODB.Connection")
		cn_DB.Open GetConnectionStr(True)
	
		Set rs_Data = Server.CreateObject("adodb.recordset")
		Set rs_Data = cn_DB.Execute("SELECT * FROM _qws_PrefInventorySync WHERE ID=" & str_ID)
		Dim str_Colour, str_ColourCode

		If Not rs_Data.EOF Then
			Response.Write("<br/>Found<br/>")
			Dim str_Part
	
			str_Colour = Replace(rs_Data("Colour"), "-", " ") & "."
			str_Part = rs_Data("Part")
			Call PrefAddTransfer(rs_Data("RecID"), rs_Data("Warehouse_In"), "ADJUSTMENT", rs_Data("Part"), rs_Data("Qty"), str_Colour, rs_Data("LengthFt"), true, "Debug", "", "", "")
		Else
			Response.Write("<br/>Not Found<br/>")
		End If

		cn_DB.Close
	End Function

	Function ReRun(str_ID)

	'Call PrefAddTransfer("71703", "HORNER", "NASHUA", "QUE-63", 1, "MILL", 18, true)
	'Call PrefAddTransfer("71703", "Test", "NASHUA", "QUE-63", 1, "MILL", 18, true)
	'Call PrefAddTransfer("71150", "DURAPAINT(WIP)", "NASHUA", "QUE-63", 7, "SJU Ext.", 18, true)

	Dim cn_DB: Set cn_DB = Server.CreateObject("ADODB.Connection")
	cn_DB.Open GetConnectionStr(True)

	Set rs_Data = Server.CreateObject("adodb.recordset")
	Set rs_Data = cn_DB.Execute("SELECT * FROM _qws_PrefInventorySync WHERE ID=" & str_ID)   '11934 2774
	Dim str_Colour, str_ColourCode

	If 1 = 2 Then
	ElseIf 1 = 1 Then
	End If

	Response.Write(GetScriptNameOnly("http://172.18.13.31:8081/stockpending.asp"))

	'Call PrefAddTransfer(61576, "NASHUA", "ADJUSTMENT", "Que-141", 2, "VRD-Ext", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(75764, "NASHUA", "ADJUSTMENT", "Que-17", 40, "K1285", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "", "NASHUA", "Que-143", 2, "UC106683F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11487, "ADJUSTMENT", "NASHUA", "NC70001", 8, "MILL", 21.33, true, "Debug", "", "")
	'Call PrefAddTransfer(11487, "NASHUA", "ADJUSTMENT", "NC70001", 8, "UC114626", 21.33, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "NC70023", 6, "CLR-ANDZD", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "ADJUSTMENT", "HORNER", "Que-141", 1, "MILL", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(11191, "HORNER", "ADJUSTMENT", "Que-61", 53, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11189, "HORNER", "ADJUSTMENT", "Que-108", 525, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11189, "HORNER", "ADJUSTMENT", "Que-108", 525, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11189, "ADJUSTMENT", "HORNER", "Que-162", 149, "MILL", 22, true, "Debug", "", "")
	'Call PrefAddTransfer(11189, "HORNER", "ADJUSTMENT", "Que-62", 580, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11226, "HORNER", "ADJUSTMENT", "Que-142", 743, "MILL", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(11207, "HORNER", "ADJUSTMENT", "Que-145", 751, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11217, "ADJUSTMENT", "HORNER", "Que-146", 325, "MILL", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(11260, "HORNER", "ADJUSTMENT", "Que-147", 418, "MILL", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "2NC7686", 25, "K1285", 21, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-128", 16, "SJX-Ext", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-151", 83, "K1285", 20, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-28", 7, "CCM-Ext", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-94", 10, "UC70470F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-95", 10, "UC70470F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "ADJUSTMENT", "NASHUA", "Que-95", 2, "UC70470F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "ADJUSTMENT", "NASHUA", "Que-146", 126, "UC70470F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-146", 126, "UC70470F", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "ADJUSTMENT", "NASHUA", "Que-146", 126, "UC114626", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "ADJUSTMENT", "NASHUA", "Que-146", 138, "UC114626", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-146", 126, "UC114626", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-128", 10, "Mill", 18, true, "Debug", "", "")
	'Call PrefAddTransfer(1, "NASHUA", "ADJUSTMENT", "Que-151", 162, "Mill", 20, true, "Debug", "", "", "")
	'Call PrefAddTransfer(13946, "NASHUA", "ADJUSTMENT", "Que-66", 540, "Clear/Anod", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(13946, "NASHUA", "ADJUSTMENT", "Que-103", 491, "K1285", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(13738, "NASHUA", "ADJUSTMENT", "Que-62", 534, "K1285", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(13878, "NASHUA", "ADJUSTMENT", "Que-65", 100, "K1285", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(13878, "NASHUA", "ADJUSTMENT", "Que-65", 5, "K1285", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(16748, "NASHUA", "ADJUSTMENT", "Que-94", 10, "UC70192F", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(16676, "DURAPAINT", "ADJUSTMENT", "QUE-50", 352, "UCFX10053", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(16676, "DURAPAINT", "ADJUSTMENT", "QUE-23", 2, "UC70149F-EXT", 18, true, "Adjustment", "", "", "")
	'Call PrefAddTransfer(16676, "ADJUSTMENT", "DURAPAINT", "QUE-50", 352, "UCFX10053", 18, true, "Adjustment", "", "", "")

	'Response.Write("<br/>" & GetColourPrefNew(cn_DB, "Que-204", "SJX-Int", "UC114626") & "<br/>")
	If Not rs_Data.EOF Then
		Response.Write("<br/>Found<br/>")
		Dim str_Part

		str_Colour = Replace(rs_Data("Colour"), "-", " ") & "."
		str_Part = rs_Data("Part")
		'str_ColourCode = Trim(GetColourQuest(cn_DB, str_Colour) & "")
		'Response.Write("<br/>" & GetColourPrefNew(cn_DB, str_Part, str_Colour, str_ColourCode) & "<br/>")

		'Response.Write("<br/>Colour:" & str_Colour & ", Part:" & str_Part & ", Colour Code:" & str_ColourCode & "<br/>")
		Call PrefAddTransfer(rs_Data("RecID"), rs_Data("Warehouse_Out"), rs_Data("Warehouse_In"), rs_Data("Part"), rs_Data("Qty"), str_Colour, rs_Data("LengthFt"), true, "Debug", "", "", "")
	Else
		Response.Write("<br/>Not Found<br/>")
	End If

	'Response.Write(GetColourQuest(cn_DB, "SFX INT.") & "<br/>")
	'Response.Write(GetColourPref(cn_DB, Replace(Replace("SFX INT.", " ", "-"), ".", ""), "UCFX12487") & "<br/>")
	'Response.Write(GetPartPref(cn_DB, "QUE-141", Replace(Replace("SFX INT.", " ", "-"), ".", ""), "UCFX12487") & "<br/>")

	'Call PrefAddTransfer("71703", "HORNER", "NASHUA", "QUE-141", 1, "MILL", 18, true)
	'Call PrefAddTransfer("71703", "NASHUA", "HORNER", "Que-141", 1, "Mill", 18, true, "Test")

	'Call PrefAddTransfer("71848", "NASHUA", "WINDOW PRODUCTION", "QUE-168", 1, "LVX Int.", 18, true, "Test")
	'Call PrefAddTransfer("71848", "NASHUA", "WINDOW PRODUCTION", "que-167", 1, "White", 20, true, "Test")

	'Call PrefAddTransfer("71848", "", "NASHUA", "NC70006", 100, "FAI Ext.", 21.33, true, "Test")
	'Call PrefAddTransfer("71848", "NASHUA", "WINDOW PRODUCTION", "NC70006", 15, "FAI Ext.", 21.33, true, "Test")

'431

	'Call PrefAddTransfer("72068", "", "WINDOW PRODUCTION", "QUE-201", 200, "Mill", 20, true, "Test")
	'Call PrefAddTransfer("72068", "WINDOW PRODUCTION", "HORNER", "QUE-201", 115, "Mill", 20, true, "Test")
	'Call PrefAddTransfer("72195", "NPREP", "WINDOW PRODUCTION", "QUE-17", 31, "LVX Int.", 20, true, "Test")
	'Call PrefAddTransfer("72195", "NASHUA", "WINDOW PRODUCTION", "Que-143", 12, "PSQ Ext.", 18, true, "Test")
	'Call PrefAddTransfer("72222", "NPREP", "WINDOW PRODUCTION", "Que-17", 38, "FAI Ext.", 20, true, "Test") ' 840
	'Call PrefAddTransfer("72204", "NASHUA", "WINDOW PRODUCTION", "Que-168", 4, "FAI Ext.", 20, true, "Test") ' 769
	'Call PrefAddTransfer("72240", "SAPA", "HORNER", "QUE-203", 132, "Mill", 20, true, "Test") ' 867
	'Call PrefAddTransfer("72242", "SAPA", "HORNER", "QUE-203", 220, "Mill", 20, true, "Test") ' 869
	'Call PrefAddTransfer("72243", "SAPA", "HORNER", "QUE-203", 220, "Mill", 20, true, "Test") ' 870
	'Call PrefAddTransfer("72244", "SAPA", "HORNER", "QUE-203", 220, "Mill", 20, true, "Test") ' 872
	'Call PrefAddTransfer("70326", "SAPA", "HORNER", "QUE-203", 220, "Mill", 20, true, "Test") ' 874
	'Call PrefAddTransfer("72287", "NASHUA", "WINDOW PRODUCTION", "QUE-108", 6, "TRI Int.", 18, true, "Test") ' 874

	cn_DB.Close
	End Function
	
	Function GetColourPrefNew(cn_DB, str_Part, str_Colour, str_ColourCode)
		Dim b_Found: b_Found = False
		Dim str_Ret
		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM [Quest].[dbo].ColorConfigurations WHERE ColorName='" & str_Colour & "'")
		If Not rs_Data.EOF Then
			Dim str_TmpColour
			str_TmpColour = rs_Data(0)
			Set rs_Data = cn_DB.Execute("SELECT COUNT(*) FROM [Quest].[dbo].Materiales WHERE Referencia='" & str_Part & " " & str_Colour & "'")
			If rs_Data(0) > 0 Then
				str_Ret = str_TmpColour
				b_Found = True
			End If
		End If

		If b_Found = False Then
			Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM [Quest].[dbo].ColorConfigurations WHERE ColorName='" & str_ColourCode & "'")
			If Not rs_Data.EOF Then
				str_Ret = rs_Data(0)
			End If
		End If

		rs_Data.Close: Set rs_Data = Nothing

		GetColourPrefNew = str_Ret
	End Function

	Function GetIDByTable(isSQLServer, i_ID, str_Table)
		Dim str_Ret
		str_Ret = ""

		If gi_Mode = c_MODE_HYBRID And isSQLServer Then
			Select Case(i_ID)
				Case 1
					str_Ret = gstr_ID1
				Case 2
					str_Ret = gstr_ID2
				Case 3
					str_Ret = gstr_ID3
				Case 4
					str_Ret = gstr_ID4
				Case 5
					str_Ret = gstr_ID5
				Case 6
					str_Ret = gstr_ID6
			End Select
		ElseIf gi_Mode = c_MODE_SQL_SERVER Then

			Dim cn_DB: Set cn_DB = Server.CreateObject("ADODB.Connection")
			cn_DB.Open gstr_DB_Pref

			Dim cmd_Data
			Set cmd_Data = Server.CreateObject("ADODB.Command")
			Set cmd_Data.ActiveConnection = cn_DB
			cmd_Data.CommandText = "qws_sp_GetNextTableID"
			cmd_Data.CommandType = &H0004

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("i_ID", 3, 4)
			cmd_Data.Parameters.Append cmd_Data.CreateParameter("str_TableName", 200, &H0001, 50, str_Table)

			str_Ret = cmd_Data("i_ID")
			cn_DB.Close
		End If

		GetIDByTable = str_Ret
	End Function

%>