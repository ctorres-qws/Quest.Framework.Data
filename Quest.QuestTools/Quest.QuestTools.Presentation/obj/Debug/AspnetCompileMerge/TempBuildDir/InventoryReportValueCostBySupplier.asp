<!--#include file="dbpath_secondary.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to update logic for checking paint price and remove default value of 0.38
	Date: November 29, 2019
	Modified By: Michelle Dungo
	Changes: Modified to add Warehouse field
	
	Date: February 4, 2020
	Modified By: Michelle Dungo
	Changes: Modified to add Milvan Warehouse
-->	
<%
Server.ScriptTimeout=4000
%>
<%

	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=QCInventoryReport.xls"
	Else
%>
<style>
	body { font-family: arial; }
	td { font-size: 13px; }
</style>
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%
	End If

Dim cn_SQL: Set cn_SQL = Server.CreateObject("adodb.connection")
Dim str_Bundles, str_Supplier, str_Debug
Dim b_Debug: b_Debug = True
Dim d_PriceWhite, d_PriceColor
Dim gstr_MissingAlumTypes, gstr_MissingPricePeriod, gstr_MissingPriceColor
Dim gstr_DebugMsg
Dim gstr_ID
d_PriceWhite = 0.14: d_PriceColor = 0.38
d_PriceWhiteDef = 0.14: d_PriceColorDef = 0.38
Dim d_PriceColorLeg
ReportName = Request.Querystring("ReportName") & ""
Country = Request.Querystring("Country")
'ReportName = "Y_INV112017S"
Dim AlumPrice
AlumPrice = 3.95


' Canada View NASHUA USA view JUPITER - for future reports
if Country ="USA" then
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'JUPITER') ORDER BY PART ASC, ID"
else
	strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP' OR WAREHOUSE = 'MILVAN') ORDER BY PART ASC, ID"
end if

'Texas report will need JUPITER
Set rs = Server.CreateObject("adodb.recordset")
If IsDebug Then
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND Part = '2NC5128' ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND ID = 4222 ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND ID = 46945 ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE ID = 57761"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE ID = 24678"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE ID = 126"
	
	'strSQL = "SELECT TOP 200 * FROM [" & ReportName & "] "
End If

'Response.Write(DBConnection2)
'Response.End()

'strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND (Colour IN('Mill','White') AND Len(ColorPO)=0) order by PART ASC"
'strSQL = "SELECT yI.* FROM [" & ReportName & "] as yI LEFT JOIN Y_MASTER as yM on yM.Part = yI.Part WHERE (yI.WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP')) order by yI.PART ASC"

DBConnection2.Close
set DBConnection2 = nothing

Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN = GetConnectionStrSecondary(False) 'method in @common.asp
DBConnection2.Open DSN


rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection2

DBOpen cn_SQL, True

'Create a Query
SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
Set RS2 = GetDisconnectedRS(SQL2, cn_SQL)
'Get a Record Set

SQL2 = "SELECT * FROM _qws_Inv_SupplierPrices order BY Period DESC"
Set rs_Prices = GetDisconnectedRS(SQL2, cn_SQL)

'Set rs_Y_Colors = GetDisconnectedRS("SELECT * FROM y_Color ORDER BY ID DESC", DBConnection)
Set rs_Y_Colors = GetDisconnectedRS("SELECT yC.*, qIPF.Price, qIPF.Period, yC.Company from y_color yC INNER JOIN _qws_Inv_PaintFamily qIPF ON qIPF.PaintFamily = yC.COMPANY ORDER BY yC.ID DESC", cn_SQL)

Dim rs_Y_InvLog
If IsDebug Then
	'Set rs_Y_InvLog = GetDisconnectedRS("SELECT Bundle, Part, Colour, Min(ModifyDate) as DateIn FROM y_invlog WHERE Warehouse IN ('NASHUA','GOREWAY','HORNER') GROUP BY Bundle, Part, Colour", DBConnection)
	Set rs_Y_InvLog = GetDisconnectedRS("SELECT Bundle, Part, Min(ModifyDate) as DateIn, ItemID  FROM y_invlog WHERE Warehouse IN ('NASHUA','GOREWAY','HORNER') GROUP BY Bundle, Part, ItemID", DBConnection)
End If

cn_SQL.Close: Set cn_SQL = Nothing

'Response.Write(GetSupplier("625674"))
'Response.Write(GetAlumPrice("913617, 913618, 913619, 913620, 913779, 913780, 913781, 913782, 913783, 913784, 913789, 913790, 913792, 913793, 913803, 913804","05/26/2015"))
'Response.Write(GetAlumPrice("A12013377","5/26/2015"))
'Response.Write("Debug:" & str_Debug & "<br/>")
'Response.Write("Supplier:" & str_Supplier & "<br/>")
'Response.End()

'Aluminum price
alumprice = AlumPrice

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<style>
	.csVal { text-align: right; }
</style>
<% 
if Country ="USA" then
	
%>
<h2> Inventory from Jupiter</h2>
<%	
else
%>
<h2> Inventory from Goreway / Nashua / Nashua Prep / Durapaint / Durapaint (WIP) / Horner / Milvan / TORBRAM (For historical purposes)</h2>
<%	
end if

%>




<h3> Using Inventory from: <%Response.write ReportName%></h3>
<% If Request("Download") <> "YES" Then %>
<!--// <h3> Aluminium Price: <%Response.write AlumPrice%></h3> //-->
<a href="InventoryReportValueCostBySupplier.asp?AlumPrice=<%response.write AlumPrice%>&ReportName=<%response.write ReportName%>&Download=YES" target="_self"><b>Download Excel Copy</a><br/>
<% End If %>
<table border='1' class='sortable' cellpadding="0" cellspacing="0">
  <tr>
<% If IsDebug Then %>
    <th>ID</th>
<% End If %>
    <th>Part</th>
    <th>Qty</th>
    <th>Length (mm)</th>
    <th>Colour</th>
	<th>KGM</th>
   <th>Bundle</th>
<% If IsDebug Then %>
    <th>&nbsp;&nbsp;&nbsp;Value&nbsp;&nbsp;</th>
<% End If %>
    <th>Value(Supplier)</th>
    <th>Date In</th>
	 <th>White Paint</th>
	  <th>Color Paint</th>
<% If IsDebug Then Response.Write("<th>Color Paint(Legacy)</th>") %>
	  <th>&nbsp;&nbsp;Warehouse&nbsp;&nbsp;</th>	
	  <th>Obsolete</th>
<% If b_Debug Then Response.Write("<th>Debug</th>") %>
    </tr>

<%
Dim bCalc, d_AlumPrice, d_Value
Dim str_Period
d_AlumPrice = 0.00: d_Value = 0.00
Dim d_ObsoleteValue: d_ObsoleteValue= 0.00
Dim d_ObsoletePaintValue: d_ObsoletePaintValue= 0.00
Dim d_ObsoleteWhiteValue: d_ObsoleteWhiteValue= 0.00
Dim b_Process
bCalc = True
rs.movefirst
Dim b_Error, b_ErrorColor
Do While Not rs.eof
	gstr_ID = rs("ID")
	b_Error = False
	b_ErrorColor = False
	str_Supplier = "": str_Debug = ""
	invpart = rs("part")
	partqty = rs("Qty")
	lmm = CDBL(rs("lmm"))
	linch = CDBL(rs("linch"))
	colour = rs("colour")
	project = rs("project")
	bundle = rs("bundle")
	kgm = rs("kgm")
	datein = rs("datein")
	str_Supplier = rs("Supplier")
	d_PriceColorLeg = 0.0
	warehouse = rs("Warehouse") & ""
	
	If UCase(Trim(str_Supplier)) = "METRA-SYSTEMS" Or UCase(Trim(str_Supplier)) = "METRASYSTEMS" Then
		str_Supplier = "METRA"
	End If
	errormsg = 0
	str_Job = Trim(Replace(Replace(UCase(rs("Colour") & ""), "INT.", ""), "EXT.", ""))

	If Len(str_Job) <> 3 Then
		str_Job = Trim(Replace(Replace(UCase(rs("JobComplete") & ""), "INT.", ""), "EXT.", ""))
	End If

	If UCase(str_Job) = "AAA" OR LEN(str_Job) > 3 Then
		str_Job = ""
	End If
	

	gstr_DebugMsg = ""

	str_Period = ""
	
	If datein <> "" Then str_Period = GetPeriod(datein)

	RS2.Filter = "Part='" & Trim(invpart) & "'"

	b_Process = True

	Dim str_AlumType: str_AlumType = "" ' Solid or Hollow
	If rs2.eof Then
		errormsg = 1
	Else
		
		If kgm = "0" Then
			kgm =0
			kgm = rs2("kgm")
		End If
		If Trim(UCase(rs2("ExtrusionType"))) & "" = "SOLID" Then str_AlumType = "_S"
		If Trim(UCase(rs2("ExtrusionType"))) & "" = "HOLLOW" Then str_AlumType = "_H"
		str_Debug = str_Debug & ",IT:" & rs2("InventoryType")
		If UCase(rs2("InventoryType") & "") = "PLASTIC" Then b_Process = False
		If UCase(rs2("InventoryType") & "") = "SHEET" Then b_Process = False
		errormsg = 0
	End If

	If UCase(invpart) = "A-201" Then
		'If Trim(UCase(rs2("ExtrusionType"))) & "" = "SOLID" Then 
		'Response.Write("Len:" & Len(rs2("ExtrusionType").Value))
		'Response.Write("Part:" & invpart & ":"  & str_AlumType & ":" & rs2("ExtrusionType"))
		'Response.End()
	End If

	If errormsg = 1 Then
		response.write invpart & " not in inventory master <BR>"
	End If

	Dim b_Obsolete: b_Obsolete = false
	'If Instr(1, "[,AHX,ALG,AMH,ANY,APX,ARC,ARD,ARX,ASD,ATM,ATX,BAL,BAT,BPC,BWY,CCT,CHT,COV,DAV,EAH,EAO,EAS,EYE,FLR,GHO,GRD,GRP,GTM,HHO,HUD,INN,MIR,MPA,MPC,MPF,MPH,PRB,PRY,PTR,RUS,RVT,RVX,RVY,SAT,SFB,SHX,SOP,SPG,STA,TAB,TCX,UAA,UAB,UAX,UAY,WLD,WNU,]", "," & str_Job & ",") > 0 Then
	'If Instr(1, "[,a,]", "," & str_Job & ",") > 0 Then
	If Instr(1, "[,ALG,AMH,ARC,ARX,BAL,BAT,BPC,CCT,DAV,EAH,EAO,EYE,FLR,GHO,GRP,GTM,HHO,HUD,MIR,MPA,MPC,MPF,MPH,PRB,PTR,RUS,SAT,SFB,STA,TAB,WNU,]", "," & str_Job & ",") > 0 Then
		b_Obsolete = True
	End If

	If b_Process Then
	If Not IsNull(PARTQTY) Then

		If kgm > 5 Then
			pricebar = partqty * kgm
			value2 = value2 + pricebar
			'this code is 5 times overvalued
			str_Debug = "ID:" & gstr_ID
			If b_Obsolete Then d_ObsoleteValue = d_ObsoleteValue + pricebar
		Else
			If invpart = "Que-157" Then
				tempvalue =  0 * (partqty * (kgm * ( lmm / 1000 )))
			Else
				d_AlumPrice = GetAlumPriceSupplier(invpart, str_Supplier, datein,bundle, str_AlumType)
				tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
				d_Cost =  d_AlumPrice * (partqty * (kgm * ( lmm / 1000 )))
				If kgm = 0 Then 
					str_Supplier = ""
					str_Debug = ""
				End If
			End If

			value = value + tempvalue
			d_Value = d_Value + d_Cost

			If b_Obsolete Then d_ObsoleteValue = d_ObsoleteValue + d_Cost
			
		End If

		If colour = "White" Then
' ********* Get White Price - S
			paintlft = (linch/12) * partqty
			'totalpaintlft = totalpaintlft + paintlft
			' ** Get White Price
			paintvalue1 = paintvalue1 + (paintlft * d_PriceWhite)
			If b_Obsolete Then d_ObsoleteWhiteValue = d_ObsoleteWhiteValue + (paintlft * d_PriceWhite)
' ********* Get White Price - E
		Else
			If colour = "Mill" Then
			Else
' ********* Get Color Price - S
				paintlft2 = (linch/12) * partqty

				If datein <> "" Then

					Dim str_PaintCompany: str_PaintCompany = ""

					rs_Y_Colors.Filter = "PROJECT='" & colour & "' AND Period=" & str_Period
					If Not rs_Y_Colors.EOF Then
						d_PriceColor = rs_Y_Colors("Price")
						str_PaintCompany = rs_Y_Colors("Company")
						str_Debug = str_Debug & ",C$:" & d_PriceColor & "(PF:" & str_PaintCompany & ")"
					Else
						b_ErrorColor = True
						'remove setting to default price for color
						d_PriceColor = 0

						rs_Y_Colors.Filter = "PROJECT='" & colour & "'"
						If Not rs_Y_Colors.EOF Then
							str_PaintCompany = rs_Y_Colors("Company")
							d_PriceColor = GetPaintPrice(str_PaintCompany, str_Period)							
						Else
							gstr_MissingPriceColor = color
						End If

						If Instr(1, "," & gstr_MissingPriceColor, str_PaintCompany & "-" & str_Period) < 1 Then
							gstr_MissingPriceColor = gstr_MissingPriceColor & ""  & "," & str_PaintCompany & "-" & str_Period
						End If

						str_Debug = str_Debug & ",C$ (PF:" & str_PaintCompany & "): Default($NC)"

					End If
				End If
' ********* Get Color Price - E

				paintvalue2 = paintvalue2 + (paintlft2 * d_PriceColor)
				d_PriceColorLeg = (paintlft2 * d_PriceColorDef)
				If b_Obsolete Then d_ObsoletePaintValue = d_ObsoletePaintValue + (paintlft2 * d_PriceColor)
			End If
		End If

		If kgm = 0 Then 'Or kgm > 5 Then 
			b_Error = False
			str_Debug = ""
			gstr_DebugMsg = ""
			b_ErrorColor = False
			str_Job = ""
		End If

	Dim b_Display: b_Display = True
	If IsDebug Then
		'b_Display = False
		'If b_Obsolete Then b_Display = True
	End If
	If b_Display Then
%>

<tr <% If b_Error Then Response.Write("style='background-color: #FFFFCC;'") %>>
<% If IsDebug Then %>
	<td><% response.write gstr_ID %></td>
<% End If %>
	<td><% response.write invpart %></td>
	<td><% response.write partqty %></td>
	<td><% response.write lmm %></td>
	<td><% response.write colour %></td>
	<td><% response.write kgm %></td>
	<td><% response.write bundle %></td>
  <!--  <td></td> -->
<% If IsDebug Then %>
	<td class="csVal">$
<%
		If kgm > 5 then
			response.write round(pricebar,2)
		Else
			response.write round(tempvalue,2)
		End If
%>
</td>
<% End If %>
	<td class="csVal">$
<%
		If kgm > 5 then
			response.write round(pricebar,2)
		Else
			response.write round(d_Cost,2)
		End If
%>
</td>
	<td><% response.write datein %></td>
<%
		Dim str_ColorPriceErr: str_ColorPriceErr = ""
		If b_Error Then 
			str_ColorPriceErr = "style='background-color: #FFCC00;'"
		End If

		If colour = "White" then
			response.write "<td class='csVal'>$" & round(paintlft * d_PriceWhite,2) & "</td><td></td>"
		Else
			If colour = "Mill" then	
				response.write "<td>&nbsp;</td><td>&nbsp;</td>"
			Else
				response.write "<td></td><td " & str_ColorPriceErr & " class='csVal'>$" & round(paintlft2 * d_PriceColor,2) & "</td>"
			End If
		End If
		If IsDebug Then
			Response.Write("<td>" & d_PriceColorLeg & "</td>")
		End If
		
		Response.Write("<td>" & warehouse & "</td>")
		Response.Write("<td>" & b_Obsolete & "</td>")
		
		If str_Job <> "" Then str_Job = "Job: " & str_Job & ","
		If b_Debug Then Response.Write("<td style='font-size: 11px;'>" & str_Job & str_Debug & AppendStr(gstr_DebugMsg, "<br />") & "</td>")

%>
    </tr>
<%
		End If
	Else
		Response.Write "QTY IN INVENTORY IS ZERO <BR>"
	End If
	End If
	rs.movenext
loop
rs.close
set rs=nothing

rs2.close
set rs2 = nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing

%> </Table> <%

'paintvalue1 = (totalpaintlft * d_PriceWhite)
'paintvalue2 = (totalpaintlft2 * d_PriceColor)

response.write "<BR><HR>"
response.write "Total kgm Material $" & FormatNumber(round(d_Value,2),,,,-1) & "<BR>"
response.write "Total perbar Material $" & FormatNumber(round(value2,2),,,,-1)  & "<BR>"
response.write "Total Obsolete $" & FormatNumber(round(d_ObsoleteValue,2),,,,-1)  & "<BR>"
response.write "<HR>"
response.write "SubTotal = $" & FormatNumber(round(value,2) + round(value2,2),,,,-1) & "<BR><BR>"
response.write "SubTotal(Supplier Cost) = $" & FormatNumber(round(d_Value,2) + round(value2,2),,,,-1) & "<BR><BR>"
response.write "Total Paint White $" & FormatNumber(paintvalue1,,,,-1) & "<BR>"
response.write "Total White Obsolete$" & FormatNumber(d_ObsoleteWhiteValue,,,,-1) & "<BR>"
response.write "Total Paint Project $" & FormatNumber(paintvalue2,,,,-1) & "<BR>"
response.write "Total Paint Obsolete$" & FormatNumber(d_ObsoletePaintValue,,,,-1) & "<BR>"
response.write "<HR>"
response.write "Grand Total = $" & FormatNumber(round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2),,,,-1) & "<br>"
response.write "Grand Total (Supplier Cost)= $" & FormatNumber(round(d_Value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2),,,,-1)

'Response.Write(str_Bundles)
Response.Write("<br/><pre>Parts Missing Aluminum Type(Hollow, Solid): " & vbcrlf & Replace(gstr_MissingAlumTypes,",", vbCrLf) & "</pre>")

Response.Write("<br/><pre>Missing Supplier Prices For Period: " & vbcrlf & Replace(gstr_MissingPricePeriod,",", vbCrLf) & "</pre>")

Response.Write("<br/><pre>Missing Price Color For Paint Family & Period: " & vbcrlf & Replace(gstr_MissingPriceColor,",", vbCrLf) & "</pre>")

%>

</body>
</html>
<%

	Function GetAlumPriceSupplier(str_Part, str_Supplier, str_DateIn, str_BundleNos, str_AlumType)
		Dim str_Ret: str_Ret = ""
		Dim a_Date: a_Date = Split(str_DateIn & "", "/")
		Dim str_Period, str_PeriodIn

		Dim a_Bundles: a_Bundles = Split(Replace(str_BundleNos & "",",","/"), "/")
		Dim str_Bundle

		If UBound(a_Bundles) >= 0 Then
			'If Left(a_Bundles(0),1) = "/" Then a_Bundles(0) = Replace(a_Bundles(0), "/", "")
			str_Bundle = a_Bundles(0)
		End If

		str_Supplier = Trim(str_Supplier & "")
		If str_Supplier = "" And str_Bundle <> "" Then

			str_Supplier = GetSupplier(str_Bundle)
			'str_Debug = str_Debug & " GetSupplier - " & str_Bundle
			'Response.Write(str_Bundle & ":" & str_Supplier)
			'Response.End()
		End If
		
		If UBound(a_Date) >= 2 Then
			str_Period = a_Date(2) & Right("0" & a_Date(0), 2)
		End If

		'If IsDebug Then
		'	If UCase(str_Supplier) = "SAPAMILL" Then
		'		rs_Y_InvLog.Filter = "Bundle='" & str_Bundle & "' AND Part='" & str_Part & "'"
		'		'rs_Y_InvLog.Filter = "ItemID='" & gstr_ID & "' AND Part='" & str_Part & "'"
		'		If Not rs_Y_InvLog.EOF Then
		'			str_PeriodIn = rs_Y_InvLog("DateIn") & ""
		'			If str_PeriodIn <> "" Then
		'				a_Date = Split(str_PeriodIn & "", "/")
		'				str_PeriodIn = a_Date(2) & Right("0" & a_Date(0), 2)
		'			End If
		'		End If
		'		'Response.Write("Bundle='" & str_Bundle & "' AND Part='" & str_Part & "'")
		'		'Response.End()
		'	End If
		'End If

		Dim b_Process: b_Process = True

		If str_Period <> "" Then
			If CInt(Left(str_Period,4)) < 2016 Then
				b_Process = False
			End If
		End If

		If b_Process = False Then
			str_Ret = CDbl(3.95)
		Else
			If str_Period <> "" Then
				
				If str_AlumType <> "" Then
					str_Ret = GetPrice(str_Supplier, str_Period, str_AlumType)
					str_Debug = str_Debug & ",AT:" & str_AlumType
				Else
					str_AlumType = "_S"
					str_Ret = GetPrice(str_Supplier, str_Period, str_AlumType)
					str_Debug = str_Debug & ",AT(Def):" & str_AlumType

					'str_Debug = str_Debug & ",AT: ?"
					'If Instr(1, "," & gstr_MissingAlumTypes, str_Part) < 1 Then
					'	gstr_MissingAlumTypes = gstr_MissingAlumTypes & "," & str_Part
					'End If
					'b_Error = true
				End If
			End If
		End If

		If str_Ret = "" Then
			str_Ret = "0"
			str_Bundles = str_Bundles & str_BundleNos & ","
		End If
		'If str_Ret <> "0" Then str_Debug = "Supplier: " & str_Supplier & ",&nbsp;Price:&nbsp;" & str_Ret & str_Debug
		'If str_Bundle <> "" Then 
		str_Debug = "ID:" & gstr_ID & ",S:" & Show(str_Supplier) & ",&nbsp;$:&nbsp;" & str_Ret & str_Debug & ",P: " + str_Period & ",AI: " + str_PeriodIn
		GetAlumPriceSupplier = CDbl(str_Ret)
	End Function

	Function Show(str_Val)
		Dim str_Ret: str_Ret = str_Val
		If Trim(str_Val) = "" Then str_Ret = "?"
			Show = str_Ret
	End Function

	Function GetSupplier(str_Bundle)
		Dim str_Ret

		If IsNumeric(str_Bundle) Then
			If Left(str_Bundle,3) = "113" Then
				'str_Ret = "SAPAMONTREAL"
				str_Ret = "HYDRO"
			ElseIf Left(str_Bundle,3) = "109" and Len(str_Bundle) = 7 Then
				'str_Ret = "SAPAMILL"
				str_Ret = "HYDRO"
			ElseIf CDbl(str_Bundle) > (958153 - 300000) AND CDbl(str_Bundle) < (958153 + 100000) Then
				'str_Ret = "SAPAMILL"
				str_Ret = "HYDRO"
			ElseIf CDbl(str_Bundle) > (560442 - 100000) AND CDbl(str_Bundle) < (560442 + 100000) Then
				str_Ret = "APEL"			
			ElseIf Len(str_Bundle) = 7 Then
				str_Ret = "CANART"
			Else
				str_Ret = ""
			End If
		Else
			If Left(str_Bundle,1) = "A" Then
				str_Ret = "EXTAL"
			ElseIf UCase(Left(Trim(str_Bundle),3)) = "MCA" Then
				str_Ret = "METRA"
			Else
				str_Ret = ""
			End If
		End If

		GetSupplier = str_Ret
	End Function

	' note this method is not being used
	Function GetAlumPrice(str_BundleNos, str_DateIn)
		Dim str_Ret: str_Ret = ""
		Dim a_Bundles: a_Bundles = Split(Replace(str_BundleNos & "",",","/"), "/")
		Dim a_Date: a_Date = Split(str_DateIn & "", "/")
		Dim str_Period

		If UBound(a_Date) >= 2 Then
			str_Period = a_Date(2) & Right("0" & a_Date(0), 2)
		End If

		Dim str_Bundle

		If UBound(a_Bundles) > 0 Then
			str_Bundle = a_Bundles(0)
		End If

		str_Supplier = ""
		'If str_BundleNos = "" Then
		'	str_Ret = "0"
		'Else
			If str_Period <> "" Then
				If IsNumeric(str_Bundle) Then
					If Left(str_Bundle,3) = "113" Then
						'str_Supplier = "SapaMontreal"
						'str_Ret = GetPrice("SapaMontreal", str_Period)
						str_Supplier = "HYDRO"
						str_Ret = GetPrice("HYDRO", str_Period)
					ElseIf  Left(str_Bundle,3) = "109" AND  Len(str_Bundle) = 7 Then
						'str_Supplier = "SapaMill"
						'str_Ret = GetPrice("SapaMill", str_Period)
						str_Supplier = "HYDRO"
						str_Ret = GetPrice("HYDRO", str_Period)
					ElseIf CDbl(str_Bundle) > (958153 - 300000) AND CDbl(str_Bundle) < (958153 + 100000) Then
						'str_Supplier = "SapaMill"
						'str_Ret = GetPrice("SapaMill", str_Period)
						str_Supplier = "HYDRO"
						str_Ret = GetPrice("HYDRO", str_Period)
					ElseIf CDbl(str_Bundle) > (560442 - 100000) AND CDbl(str_Bundle) < (560442 + 100000) Then
						str_Supplier = "Apel"
						str_Ret = GetPrice("Apel", str_Period)
					ElseIf Len(str_Bundle) = 7 Then
						str_Supplier = "CanArt"
						str_Ret = GetPrice("CanArt", str_Period)
					Else
						str_Ret = GetPrice("", str_Period)
						str_Supplier = ""
					End If
				Else
					If Left(str_Bundle,1) = "A" Then
						str_Supplier = "Extal"
						str_Ret = GetPrice("Extal", str_Period)
					ElseIf UCase(Left(Trim(str_Bundle),3)) = "MCA" Then
						str_Supplier = "Metra"
						str_Ret = GetPrice("Metra", str_Period)
					Else
						str_Supplier = ""
						str_Ret = GetPrice("", str_Period)
					End If
				End If
			End If
		'End If
		If str_Ret = "" Then
			str_Ret = "0"
			str_Bundles = str_Bundles & str_BundleNos & ","
		End If
		'If str_Ret <> "0" Then str_Debug = "Supplier: " & str_Supplier & ",&nbsp;Price:&nbsp;" & str_Ret & str_Debug
		If str_Bundle <> "" Then str_Debug = "Supplier: " & str_Supplier & ",&nbsp;Price:&nbsp;" & str_Ret & str_Debug
		GetAlumPrice = CDbl(str_Ret)
	End Function

	Function GetPrice(str_Supplier, str_Period, str_AlumType)
		Dim str_Ret: str_Ret = "0"

		rs_Prices.Filter = "Period=" & str_Period
		If Not rs_Prices.EOF Then

			str_Supplier = Replace(str_Supplier & "","-","")

			If str_Supplier & "" <> "" AND str_Supplier <> "KEYMARK" AND str_Supplier <> "SAPA" Then
				'On Error Resume Next
				str_Ret = rs_Prices.Fields(str_Supplier & str_AlumType)
				'If Err.Number > 0 Then
					'Response.Write("ID:" & gstr_ID & ", Supplier:" & str_Supplier)
					'Response.End()
				'End If
			Else
				str_Ret = rs_Prices.Fields("Default" & str_AlumType)
				If str_Supplier & "" = "" Then
					str_Debug = str_Debug & "$:" & str_Ret
					gstr_DebugMsg = gstr_DebugMsg & "Msg: No Supplier Using Default Price:&nbsp;" & str_Ret
				Else
					str_Debug = str_Debug & "$:" & str_Ret
					gstr_DebugMsg = gstr_DebugMsg & "Msg: Using Default Price:&nbsp;"
				End If
			End If
		Else
			b_Error = True
			str_Debug = str_Debug & "&nbsp;"

			gstr_DebugMsg = gstr_DebugMsg & ",Msg: No Price For Period: " & str_Period 

			If Instr(1, "," & gstr_MissingPricePeriod, str_Period) < 1 Then
				gstr_MissingPricePeriod = gstr_MissingPricePeriod & "," & str_Period
			End If
		End If
		
		GetPrice = str_Ret
	End Function

	Function GetPeriod(str_DateIn)
		Dim str_Ret: str_Ret = ""
		Dim a_Date: a_Date = Split(str_DateIn & "", "/")

		If UBound(a_Date) >= 2 Then
			str_Ret = a_Date(2) & Right("0" & a_Date(0), 2)
		End If
		GetPeriod = str_Ret
	End Function

	Function AppendStr(str_Val, str_Sep)
		Dim str_Ret
		If str_Val <> "" Then str_Ret = str_Sep
		str_Ret = str_Ret & str_Val 
		AppendStr = str_Ret
	End Function
	
	Function GetPaintPrice(str_PaintCompany, str_Period)
		Dim str_Ret: str_Ret = "0"
		
		'get nearest price in table if item doesn't have price for that period

			'filter by masterid
			rs_Y_Colors.Filter = "Company='" & str_PaintCompany & "'"
			
			'if any price exists
			If Not rs_Y_Colors.EOF Then
				pricePeriodOlder = ""
				pricePeriodNewer = ""
				pricePeriodCurrent = ""				
				'sort ascending by period
				rs_Y_Colors.Sort = "Period ASC" 
				'convert entry date period to Long 
				entryPeriod = CLng(str_Period)
				Do While Not rs_Y_Colors.EOF
					'convert current price table period to Long 
					pricePeriodCurrent = CLng(rs_Y_Colors.Fields("Period"))					
					
					'sets value to nearest period available less than entry date (old price)
					If entryPeriod > pricePeriodCurrent Then
						pricePeriodOlder = pricePeriodCurrent
					ElseIf entryPeriod < pricePeriodCurrent Then
						'sets value to  the first period available greater than the entry date (future price)
						If pricePeriodNewer = "" Then
							pricePeriodNewer = pricePeriodCurrent
						End If
					End If
					rs_Y_Colors.MoveNext
				Loop
				
				'always use older price if available, else use future price
				If pricePeriodOlder <> "" Then
					pricePeriod = pricePeriodOlder
				Else
					pricePeriod = pricePeriodNewer
				End If
				
				' filter given masterid and nearest available price period
				rs_Y_Colors.Filter = "Company='" & str_PaintCompany & "' AND Period=" & pricePeriod
				
				'provided amount is in CAD
				If Country ="USA" Then
					str_Ret = rs_Y_Colors.Fields("Price") / GetExchangeRate(str_Period)
				Else
					str_Ret = rs_Y_Colors.Fields("Price")
				End If		
				
			'if no price available
			Else
				b_Error = True
				pricePeriod = "" ' set to blank period
			End If
		
		GetPaintPrice = str_Ret
	End Function	

%>