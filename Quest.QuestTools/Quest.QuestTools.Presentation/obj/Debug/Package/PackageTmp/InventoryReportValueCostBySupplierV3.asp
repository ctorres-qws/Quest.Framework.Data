<!--#include file="dbpath_secondary.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
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
ReportName = Request.Querystring("ReportName") & ""
'ReportName = "Y_INV112017S"
Dim AlumPrice
AlumPrice = 3.95

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') ORDER BY PART ASC, ID"
If IsDebug Then
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') AND Part = '2NC5128' ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') AND ID = 4222 ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') AND ID = 46945 ORDER BY PART ASC"
	'strSQL = "SELECT * FROM [" & ReportName & "] WHERE ID = 57761"
	'strSQL = "SELECT TOP 200 * FROM [" & ReportName & "] WHERE WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP') ORDER BY PART ASC, ID"
End If

'Response.Write(DBConnection2)
'Response.End()

'strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND (Colour IN('Mill','White') AND Len(ColorPO)=0) order by PART ASC"
'strSQL = "SELECT yI.* FROM [" & ReportName & "] as yI LEFT JOIN Y_MASTER as yM on yM.Part = yI.Part WHERE (yI.WAREHOUSE IN ('GOREWAY','NASHUA','TORBRAM','DURAPAINT','DURAPAINT(WIP)','HORNER','NPREP')) order by yI.PART ASC"

DBConnection2.Close

DSN = GetConnectionStrSecondary(False) 'method in @common.asp
DBConnection2.Open DSN


rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection2

DBOpen cn_SQL, True

'Create a Query
'SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
'Set RS2 = GetDisconnectedRS(SQL2, cn_SQL)
'Get a Record Set

'SQL2 = "SELECT * FROM _qws_Inv_SupplierPrices order BY Period DESC"
'Set rs_Prices = GetDisconnectedRS(SQL2, cn_SQL)

'Set rs_Y_Colors = GetDisconnectedRS("SELECT * FROM y_Color ORDER BY ID DESC", DBConnection)
'Set rs_Y_Colors = GetDisconnectedRS("SELECT yC.*, qIPF.Price, qIPF.Period, yC.Company from y_color yC INNER JOIN _qws_Inv_PaintFamily qIPF ON qIPF.PaintFamily = yC.COMPANY ORDER BY yC.ID DESC", cn_SQL)

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
<h2> Inventory from Goreway / Nashua / Nashua Prep / Durapaint / Durapaint (WIP) / Horner / TORBRAM (For historical purposes)</h2>
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
	  <th>Obsolete</th>
<%
	If b_Debug Then 
		Response.Write("<th>Debug</th>")
		Response.Write("<th>KGM</th>")
		Response.Write("<th>C$</th>")
		Response.Write("<th>Coords</th>")
	End If
%>
    </tr>

<%
Dim bCalc, d_AlumPrice, d_Value
Dim str_Period
d_AlumPrice = 0.00: d_Value = 0.00
Dim d_ObsoleteValue: d_ObsoleteValue= 0.00
Dim d_ObsoletePaintValue: d_ObsoletePaintValue= 0.00
Dim d_ObsoleteWhiteValue: d_ObsoleteWhiteValue= 0.00
Dim str_Coords
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
	lmm = rs("lmm")
	linch = rs("linch")
	colour = rs("colour")
	project = rs("project")
	bundle = rs("bundle")
	kgm = rs("kgm")
	datein = rs("datein")
	str_Coords = ""
	str_Supplier = rs("Supplier")
	errormsg = 0
	d_AlumPrice = 0.00
	d_PriceColor = 0.00
	str_DebugInfo = rs("Debug")

	pricebar = 0.0: d_Cost = 0.0

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

'	RS2.Filter = "Part='" & invpart & "'"

	If errormsg = 1 Then
		response.write invpart & " not in inventory master <BR>"
	End If

	Dim b_Obsolete: b_Obsolete = false
	'If Instr(1, "[,AHX,ALG,AMH,ANY,APX,ARC,ARD,ARX,ASD,ATM,ATX,BAL,BAT,BPC,BWY,CCT,CHT,COV,DAV,EAH,EAO,EAS,EYE,FLR,GHO,GRD,GRP,GTM,HHO,HUD,INN,MIR,MPA,MPC,MPF,MPH,PRB,PRY,PTR,RUS,RVT,RVX,RVY,SAT,SFB,SHX,SOP,SPG,STA,TAB,TCX,UAA,UAB,UAX,UAY,WLD,WNU,]", "," & str_Job & ",") > 0 Then
	'If Instr(1, "[,a,]", "," & str_Job & ",") > 0 Then
	If Instr(1, "[,ALG,AMH,ARC,ARX,BAL,BAT,BPC,CCT,DAV,EAH,EAO,EYE,FLR,GHO,GRP,GTM,HHO,HUD,MIR,MPA,MPC,MPF,MPH,PRB,PTR,RUS,SAT,SFB,STA,TAB,WNU,]", "," & str_Job & ",") > 0 Then
		b_Obsolete = True
	End If

	str_Coords = rs("Coords")

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
				d_AlumPrice = rs("PriceAlum") & "" 'GetAlumPriceSupplier(invpart, str_Supplier,datein,bundle,str_AlumType)
				If d_AlumPrice = "" Then d_AlumPrice = 0
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
				'totalpaintlft2 = totalpaintlft2 + paintlft2

				Dim str_PaintCompany: str_PaintCompany = ""

				If datein <> "" And rs("PriceColour") & "" <> "" Then
					d_PriceColor = rs("PriceColour")
				End If
' ********* Get Color Price - E

				paintvalue2 = paintvalue2 + (paintlft2 * d_PriceColor)
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
	<td>
<% If IsDebug Then %>
<% Else %>
		<% response.write bundle %>
<% End If %>
	</td>
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
		Response.Write("<td>" & b_Obsolete & "</td>")

		If str_Job <> "" Then str_Job = "Job: " & str_Job & ","
		If IsDebug Then 
			'Response.Write("<td style='font-size: 11px;'>" & str_Job & str_Debug & AppendStr(gstr_DebugMsg, "<br />") & "</td>")
			Response.Write("<td style='font-size: 11px;'>" & str_DebugInfo & "</td>")

			If d_AlumPrice > 0 Then 
				Response.Write("<td>" & round(d_AlumPrice, 2) & "</td>")
			Else 
				Response.Write("<td></td>") 
			End If

			If d_PriceColor > 0 Then 
				Response.Write("<td>" & round(d_PriceColor, 2) & "</td>") 
			Else 
				Response.Write("<td></td>") 
			End If

			Response.Write("<td>" & str_Coords & "</td>")
		End If

%>
    </tr>
<%
		End If
	Else
		Response.Write "QTY IN INVENTORY IS ZERO <BR>"
	End If

	rs.movenext
loop
rs.close
set rs=nothing

'rs2.close
'set rs2 = nothing

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
If IsDebug Then response.write "SubTotal = $" & FormatNumber(round(value,2) + round(value2,2),,,,-1) & "<BR><BR>"
response.write "SubTotal(Supplier Cost) = $" & FormatNumber(round(d_Value,2) + round(value2,2),,,,-1) & "<BR><BR>"
response.write "Total Paint White $" & FormatNumber(paintvalue1,,,,-1) & "<BR>"
response.write "Total White Obsolete$" & FormatNumber(d_ObsoleteWhiteValue,,,,-1) & "<BR>"
response.write "Total Paint Project $" & FormatNumber(paintvalue2,,,,-1) & "<BR>"
response.write "Total Paint Obsolete$" & FormatNumber(d_ObsoletePaintValue,,,,-1) & "<BR>"
response.write "<HR>"
If IsDebug Then response.write "Grand Total = $" & FormatNumber(round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2),,,,-1) & "<br>"
response.write "Grand Total (Supplier Cost)= $" & FormatNumber(round(d_Value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2),,,,-1)

'Response.Write(str_Bundles)
Response.Write("<br/><pre>Parts Missing Aluminum Type(Hollow, Solid): " & vbcrlf & Replace(gstr_MissingAlumTypes,",", vbCrLf) & "</pre>")

Response.Write("<br/><pre>Missing Supplier Prices For Period: " & vbcrlf & Replace(gstr_MissingPricePeriod,",", vbCrLf) & "</pre>")

Response.Write("<br/><pre>Missing Price Color For Paint Family & Period: " & vbcrlf & Replace(gstr_MissingPriceColor,",", vbCrLf) & "</pre>")

%>

</body>
</html>
<%

	Function GetSupplier(str_Bundle)
		Dim str_Ret

		If IsNumeric(str_Bundle) Then
			If Left(str_Bundle,3) = "113" Then
				str_Ret = "SAPAMONTREAL"
			ElseIf Len(str_Bundle) = 7 Then
				str_Ret = "CANART"
			ElseIf CDbl(str_Bundle) > (958153 - 300000) AND CDbl(str_Bundle) < (958153 + 100000) Then
				str_Ret = "SAPAMILL"
			ElseIf CDbl(str_Bundle) > (560442 - 100000) AND CDbl(str_Bundle) < (560442 + 100000) Then
				str_Ret = "APEL"
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

%>