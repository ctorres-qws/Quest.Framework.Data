<!--#include file="dbpath_secondary.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!--Date: August 30, 2019
	Modified By: Michelle Dungo
	Changes: Modified to update logic for checking paint price and remove default value of 0.38
-->
<%
Server.ScriptTimeout=4000
%>
<%

	If Request("Download") = "YES" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-Disposition", "attachment; filename=QCInventoryReportUS.xls"
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

' variable declaration
Dim cn_SQL: Set cn_SQL = Server.CreateObject("adodb.connection")
' comment for testing
ReportName = Request.Querystring("ReportName") & ""
'ReportName = "Y_INV042019_TEST"
Dim AlumPrice
AlumPrice = 3.95
Dim b_Debug: b_Debug = False

strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'JUPITER') ORDER BY PART ASC, ID" 'gets values from Y_INV snapshot

Set rs = Server.CreateObject("adodb.recordset")

DBConnection2.Close

' connects to MsAccess DB: InventoryReports.mdb
DSN = GetConnectionStrSecondary(False) 'method in @common.asp
DBConnection2.Open DSN

rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection2

' connect to lassard
DBOpen cn_SQL, True

SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
' get values from lassard
Set rs2 = GetDisconnectedRS(SQL2, cn_SQL)

SQL2 = "SELECT * FROM _qws_Inv_SupplierPrices order BY Period DESC"
' get values from lassard
Set rs_Prices = GetDisconnectedRS(SQL2, cn_SQL)

' get values from lassard
Set rs_Y_Colors = GetDisconnectedRS("SELECT yC.*, qIPF.Price, qIPF.Period, yC.Company from y_color yC INNER JOIN _qws_Inv_PaintFamily qIPF ON qIPF.PaintFamily = yC.COMPANY ORDER BY yC.ID DESC", cn_SQL)

cn_SQL.Close: Set cn_SQL = Nothing

'Aluminum price
alumprice = AlumPrice

%> 

<style>
	.csVal { text-align: right; }
</style>

<h2> Inventory from Jupiter</h2>

<h3> Using Inventory from: <%Response.write ReportName%></h3>
<% If Request("Download") <> "YES" Then %>
<a href="InventoryReportValueCostBySupplier_US.asp?ReportName=<%response.write ReportName%>&Download=YES" target="_self"><b>Download Excel Copy</a><br/>
<% End If %>
<table border='1' class='sortable' cellpadding="0" cellspacing="0">
  <tr>
<% If b_Debug Then %>
    <th>ID</th>
<% End If %>
    <th>Part</th>
    <th>&nbsp;&nbsp;Qty&nbsp;&nbsp;</th>
    <th>&nbsp;&nbsp;Length (in)&nbsp;&nbsp;</th>
    <th>&nbsp;&nbsp;Colour&nbsp;&nbsp;</th>
<% If b_Debug Then %>		
	<th>Kgm</th>
	<th>&nbsp;Price Per Unit (CAD)&nbsp;</th>	
<% End If %>	
	<th>&nbsp;&nbsp;Lb/Ft&nbsp;&nbsp;</th>
	<th>&nbsp;Price Per Unit (USD)&nbsp;</th>	
   <th>&nbsp;&nbsp;Bundle&nbsp;&nbsp;</th>
<% If b_Debug Then %>
    <th>&nbsp;&nbsp;&nbsp;Value (CAD)&nbsp;&nbsp;</th>
<% End If %>
    <th>&nbsp;&nbsp;&nbsp;Value(Supplier)&nbsp;&nbsp;</th>
    <th>Date In</th>
<% If b_Debug Then %>	
	 <th>White Paint (CAD)</th>
	  <th>Color Paint (CAD)</th>
<% End If %>	  
	 <th>&nbsp;&nbsp;White Paint&nbsp;&nbsp;</th>
	  <th>&nbsp;&nbsp;Color Paint&nbsp;&nbsp;</th>	  
<% If IsDebug Then Response.Write("<th>Color Paint(Legacy)</th>") %>
	  <th>Obsolete</th>
<% If b_Debug Then Response.Write("<th>Debug</th>") %>
    </tr>

<%
Dim str_Bundles, str_Supplier, str_Debug
Dim d_PriceWhite, d_PriceColor, d_PriceWhite_USD, d_PriceColor_USD
Dim gstr_MissingPricePeriod, gstr_MissingPriceColor
Dim gstr_DebugMsg
Dim gstr_ID
Dim d_AlumPrice, d_Value, d_Value_USD
Dim str_Period
Dim d_ObsoleteValue
Dim d_ObsoletePaintValue
Dim d_ObsoleteWhiteValue
Dim b_Process
Dim b_Error
Dim lbft

' set defaults
d_ObsoleteValue= 0.00: d_ObsoletePaintValue= 0.00: d_ObsoleteWhiteValue= 0.00
d_AlumPrice = 0.00: d_Value = 0.00: d_Value_USD = 0.00
d_PriceWhite = 0.14: d_PriceColor = 0.38
d_PriceWhiteDef = 0.14: d_PriceColorDef = 0.38

rs.movefirst
Do While Not rs.eof
	gstr_ID = rs("ID")
	b_Error = False
	str_Supplier = "": str_Debug = ""
	invpart = rs("part")
	partqty = rs("Qty")
	lmm = CDBL(rs("lmm"))
	linch = CDBL(rs("linch"))
	colour = rs("colour") & ""
	project = rs("project")
	bundle = rs("bundle")
	kgm = rs("kgm")
	datein = rs("datein")
	str_Supplier = rs("Supplier")
	jobComplete = rs("JobComplete") & ""
	po = rs("PO") & ""

	errormsg = 0
	gstr_DebugMsg = ""
	
	' get Job value
	str_Job = ""
	If colour <> "" Or jobComplete <> "" Then str_Job = GetJob(colour,jobComplete)
	
' start - skip to next record if test
If Instr(1, "[,AAA,]", "," & str_Job & ",") = 0 Then
	' get Period value
	str_Period = ""
	If datein <> "" Then str_Period = GetPeriod(datein)
	
	' hardcode supplier to HYDRO (previously SAPAMILL) 
	str_Supplier = "HYDRO"
	
	' Set supplier name to SAPAMILL if supplier contains the text 'SAPA'
	'If str_supplier If Instr(1, str_supplier & "","SAPA", vbTextCompare) > 0 Then
	'	str_Supplier = "SAPAMILL"
	'End if

	' filters records from Y_MASTER using the Part
	rs2.Filter = "Part='" & Trim(invpart) & "'"

	' sets default price to 3.95 if FALSE, sets to false if period year < 2016 or inventory type is PLASTIC or SHEET
	b_Process = True

	' check type of material
	Dim str_AlumType: str_AlumType = "" ' Solid or Hollow
	If rs2.eof Then
		errormsg = 1
	Else
		
		' if kgm from msAccess is 0 get value from Y_MASTER
		If kgm = "0" Then
			kgm = 0
			kgm = rs2("kgm")			
		End If
		
		If Trim(UCase(rs2("ExtrusionType"))) & "" = "SOLID" Then str_AlumType = "_S"
		If Trim(UCase(rs2("ExtrusionType"))) & "" = "HOLLOW" Then str_AlumType = "_H"
		str_Debug = str_Debug & ",IT:" & rs2("InventoryType")
		If UCase(rs2("InventoryType") & "") = "PLASTIC" Then b_Process = False
		If UCase(rs2("InventoryType") & "") = "SHEET" Then b_Process = False
		errormsg = 0
	End If	

	If errormsg = 1 Then
		response.write invpart & " not in inventory master <BR>"
	End If

	' identify obsolete jobs
	Dim b_Obsolete: b_Obsolete = false
	If Instr(1, "[,ALG,AMH,ARC,ARX,BAL,BAT,BPC,CCT,DAV,EAH,EAO,EYE,FLR,GHO,GRP,GTM,HHO,HUD,MIR,MPA,MPC,MPF,MPH,PRB,PTR,RUS,SAT,SFB,STA,TAB,WNU,]", "," & str_Job & ",") > 0 Then
		b_Obsolete = True
	End If

	If b_Process Then
	If Not IsNull(partqty) Then

		If kgm > 5 Then
		
			' for testing CAD Value (Supplier) 
			pricebar = partqty * kgm
			value2 = value2 + pricebar
			
			' for kgm > 5 kgm is price per unit
			unitPrice_CAD = kgm
			
			' convert unit price to USD
			unitPrice_USD = unitPrice_CAD / GetExchangeRate(str_Period)			
			lbft = unitPrice_USD
			
			'transformationCost_USD = CalculateTransformationCost (unitPrice_CAD,str_Period)
			
			'Value (Supplier in USD), (lbft * QTY) 		
			pricebar_USD = (lbft * partqty )
			
			value2_USD = value2_USD + pricebar_USD
			
			str_Debug = "ID:" & gstr_ID
			If b_Obsolete Then d_ObsoleteValue = d_ObsoleteValue + pricebar_USD
		
		Else
			' convert kgm to lb/ft
			lbft = ConvertToLbFt(kgm,lmm,linch)

			If invpart = "Que-157" Then
				tempvalue =  0
				tempvalue_USD = 0
			Else
			
				' for testing CAD Value (Supplier) 
				d_AlumPrice = GetAlumPriceSupplier(invpart, str_Supplier, datein,bundle, str_AlumType)
				tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
				d_Cost =  d_AlumPrice * (partqty * (kgm * ( lmm / 1000 )))
				
				' get supplier price in CAD
				alumPrice_CAD = GetAlumPriceSupplierUS(invpart, str_Supplier, datein,bundle, str_AlumType)
				'alumPrice_CAD = alumPrice_CAD / 2.20462 ' convert price per kg to price per lb
				
				
				' convert supplier price to USD
				alumPrice_USD = alumPrice_CAD / (GetExchangeRate(str_Period)*(2.20462))
				
				'transformationCost_USD = CalculateTransformationCost (alumPrice_CAD,str_Period)
			
				'Value (Supplier in USD), (alumPrice_USD * QTY * lb/ft * length (ft)) 		
				d_Cost_USD = (alumPrice_USD * (partqty * (lbft * ( linch / 12 )) ))
				'd_Cost_USD = alumPrice_USD * partqty * lbft * ( linch / 12 )
				
				' Subtotal is computed using default price of 3.95 ?
				'transformationCostTemp_USD = (transactionprice_CAD - ALUMPRICE) / GetExchangeRate(str_Period)
				
				' Subtotal is computed using default price of 3.95
				tempvalue_USD =  ((ALUMPRICE / (GetExchangeRate(str_Period)*(2.20462))) * (partqty * (lbft * ( linch / 12 )) ))
				
				If kgm = 0 Then 
					str_Supplier = ""
					str_Debug = ""
				End If
				
			End If

			value = value + tempvalue
			d_Value = d_Value + d_Cost
			
			value_USD = value_USD + tempvalue_USD
			d_Value_USD = d_Value_USD + d_Cost_USD
			str_Debug = str_Debug & ",LBFT:" & lbft

			If b_Obsolete Then d_ObsoleteValue = d_ObsoleteValue + d_Cost_USD
			
		End If
		
		' get paint price conversions from CAD to USD
		d_PriceWhite_USD = d_PriceWhite / GetExchangeRate(str_Period) 
		d_PriceColor_USD = d_PriceColor / GetExchangeRate(str_Period) 
		d_PriceWhiteDef_USD = d_PriceWhiteDef / GetExchangeRate(str_Period) 
		d_PriceColorDef_USD = d_PriceColorDef / GetExchangeRate(str_Period) 

		If colour = "White" Then
' ********* Get White Price - S
			paintlft = (linch/12) * partqty

			' ** Get White Price
			paintvalue1 = paintvalue1 + (paintlft * d_PriceWhite)
			paintvalue1_USD = paintvalue1_USD + (paintlft * d_PriceWhite_USD)
			If b_Obsolete Then d_ObsoleteWhiteValue = d_ObsoleteWhiteValue + (paintlft * d_PriceWhite_USD)
' ********* Get White Price - E
		Else
			If colour = "Mill" Then
				' do nothing
			Else
' ********* Get Color Price - S
				paintlft2 = (linch/12) * partqty

				If datein <> "" Then

					Dim str_PaintCompany: str_PaintCompany = ""

					rs_Y_Colors.Filter = "PROJECT='" & colour & "' AND Period=" & str_Period
					If Not rs_Y_Colors.EOF Then
						d_PriceColor = rs_Y_Colors("Price")
						d_PriceColor_USD = rs_Y_Colors("Price") / GetExchangeRate(str_Period) 
						str_PaintCompany = rs_Y_Colors("Company")
						str_Debug = str_Debug & ",C$:" & d_PriceColor & "(PF:" & str_PaintCompany & ")"
					Else
						d_PriceColor = d_PriceColorDef
						'remove setting to default price for color
						d_PriceColor_USD = 0

						rs_Y_Colors.Filter = "PROJECT='" & colour & "'"
						If Not rs_Y_Colors.EOF Then
							str_PaintCompany = rs_Y_Colors("Company")
							d_PriceColor_USD = GetPaintPrice(str_PaintCompany, str_Period)	'conversion is within the function
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
				paintvalue2_USD = paintvalue2_USD + (paintlft2 * d_PriceColor_USD)
				
				If b_Obsolete Then d_ObsoletePaintValue = d_ObsoletePaintValue + (paintlft2 * d_PriceColor_USD)
			End If
		End If

		If kgm = 0 Then 'Or kgm > 5 Then 
			b_Error = False
			str_Debug = ""
			gstr_DebugMsg = ""
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
<% If b_Debug Then %>
	<td><% response.write gstr_ID %></td>
<% End If %>
	<td><% response.write invpart %></td>
	<td><% response.write partqty %></td>
	<td><% response.write linch %></td>
	<td><% response.write colour %></td>
<% If b_Debug Then %>	
<%		If kgm > 5 then
			response.write "<td></td>"
			response.write "<td class=""csVal"">$"&round(kgm,2)&"</td>"
		Else			
			response.write "<td>"&round(kgm,2)&"</td>"
			response.write "<td></td>"
		End If	
%>
<% End If %>	
<%		If kgm > 5 then
			response.write "<td></td>"
			response.write "<td class=""csVal"">$"&round(lbft,2)&"</td>"
		Else			
			response.write "<td>"&round(lbft,2)&"</td>"
			response.write "<td></td>"
		End If	
%>
	<td><% response.write bundle %></td>
  <!--  <td></td> -->
<% If b_Debug Then %>
	<td class="csVal">$
<%
		If kgm > 5 then
			response.write round(pricebar,2)
		Else
			response.write round(d_Cost,2)
		End If
%>
</td>
<% End If %>
	<td class="csVal">$
<%
		If kgm > 5 then
			response.write round(pricebar_USD,2)
		Else
			response.write round(d_Cost_USD,2)
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
			If b_Debug Then
				response.write "<td class='csVal'>$" & round(paintlft * d_PriceWhite,2) & "</td>"
				response.write "<td class='csVal'>&nbsp;$0</td>"
			End If			
			response.write "<td class='csVal'>$" & round(paintlft * d_PriceWhite_USD,2) & "</td>"
			response.write "<td class='csVal'>&nbsp;$0</td>"
		Else
			If colour = "Mill" then	
				If b_Debug Then
					response.write "<td class='csVal'>&nbsp;$0</td><td class='csVal'>&nbsp;$0</td>"
				End If			
				response.write "<td class='csVal'>&nbsp;$0</td><td class='csVal'>&nbsp;$0</td>"
			Else
				If b_Debug Then
					response.write "<td class='csVal'>&nbsp;$0</td>"				
					response.write "<td " & str_ColorPriceErr & " class='csVal'>$" & round(paintlft2 * d_PriceColor,2) & "</td>"
				End If
				response.write "<td class='csVal'>&nbsp;$0</td>"
				response.write "<td " & str_ColorPriceErr & " class='csVal'>$" & round(paintlft2 * d_PriceColor_USD,2) & "</td>"
			End If
		End If

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
End If ' end check actual data
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
response.write "Total lbft Material $" & FormatNumber(round(d_Value_USD,2),,,,-1) & "<BR>"
response.write "Total perbar Material $" & FormatNumber(round(value2_USD,2),,,,-1)  & "<BR>"
response.write "Total Obsolete $" & FormatNumber(round(d_ObsoleteValue,2),,,,-1)  & "<BR>"
response.write "<HR>"
response.write "SubTotal(3.95 Alum Price) = $" & FormatNumber(round(value_USD,2) + round(value2_USD,2),,,,-1) & "<BR><BR>"
response.write "SubTotal(Supplier Cost) = $" & FormatNumber(round(d_Value_USD,2) + round(value2_USD,2),,,,-1) & "<BR><BR>"
response.write "Total Paint White $" & FormatNumber(paintvalue1_USD,,,,-1) & "<BR>"
response.write "Total White Obsolete$" & FormatNumber(d_ObsoleteWhiteValue,,,,-1) & "<BR>"
response.write "Total Paint Project $" & FormatNumber(paintvalue2_USD,,,,-1) & "<BR>"
response.write "Total Paint Obsolete$" & FormatNumber(d_ObsoletePaintValue,,,,-1) & "<BR>"
response.write "<HR>"
response.write "Grand Total(3.95 Alum Price) = $" & FormatNumber(round(value_USD,2) + round(value2_USD,2) + round(paintvalue1_USD,2) + round(paintvalue2_USD,2),,,,-1) & "<br>"
response.write "Grand Total(Supplier Cost)= $" & FormatNumber(round(d_Value_USD,2) + round(value2_USD,2) + round(paintvalue1_USD,2) + round(paintvalue2_USD,2),,,,-1)

response.Write("<br/><pre>Missing Supplier Prices For Period: " & vbcrlf & Replace(gstr_MissingPricePeriod,",", vbCrLf) & "</pre>")

response.Write("<br/><pre>Missing Price Color For Paint Family & Period: " & vbcrlf & Replace(gstr_MissingPriceColor,",", vbCrLf) & "</pre>")

%>

</body>
</html>
<%

	Function GetJob(str_Colour, str_JobComplete)
		Dim str_Ret
		
		' get str_Job value from Colour if length = 3, else get from JobComplete column			
		str_Ret = Trim(Replace(Replace(UCase(str_Colour), "INT.", ""), "EXT.", ""))

		If Len(str_Ret) <> 3 Then
			str_Ret = Trim(Replace(Replace(UCase(str_JobComplete), "INT.", ""), "EXT.", ""))
		End If

		' if its a test record or not a valid job code, set to empty string
		If UCase(str_Ret) = "AAA" OR Len(str_Ret) > 3 Then
			str_Ret = ""
		End If

		GetJob = str_Ret
	End Function

	Function GetAlumPriceSupplier(str_Part, str_Supplier, str_DateIn, str_BundleNos, str_AlumType)
		Dim str_Ret: str_Ret = ""
		Dim a_Date: a_Date = Split(str_DateIn & "", "/")
		Dim str_Period, str_PeriodIn

		Dim a_Bundles: a_Bundles = Split(Replace(str_BundleNos & "",",","/"), "/")
		Dim str_Bundle

		If UBound(a_Bundles) >= 0 Then
			str_Bundle = a_Bundles(0)
		End If

		str_Supplier = Trim(str_Supplier & "")
		
		If UBound(a_Date) >= 2 Then
			str_Period = a_Date(2) & Right("0" & a_Date(0), 2)
		End If

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
				Else
					str_AlumType = "_S"
					str_Ret = GetPrice(str_Supplier, str_Period, str_AlumType)
				End If
			End If
		End If

		If str_Ret = "" Then
			str_Ret = "0"
		End If
		GetAlumPriceSupplier = CDbl(str_Ret)
	End Function

	Function Show(str_Val)
		Dim str_Ret: str_Ret = str_Val
		If Trim(str_Val) = "" Then str_Ret = "?"
			Show = str_Ret
	End Function

	Function GetPrice(str_Supplier, str_Period, str_AlumType)
		Dim str_Ret: str_Ret = "0"

		rs_Prices.Filter = "Period=" & str_Period
		If Not rs_Prices.EOF Then

			str_Supplier = Replace(str_Supplier & "","-","")

			If str_Supplier & "" <> "" AND str_Supplier <> "KEYMARK" AND str_Supplier <> "SAPA" Then
				str_Ret = rs_Prices.Fields(str_Supplier & str_AlumType)
			Else
				str_Ret = rs_Prices.Fields("Default" & str_AlumType)
				If str_Supplier & "" = "" Then
				Else
				End If
			End If
		Else
			b_Error = True
		End If
		
		GetPrice = str_Ret
	End Function

	' get Exchange Rate from USD to CAD given the datein period, multiply to USD to get CAD, divide CAD with this value to get USD
	Function GetExchangeRate(str_Period)
		Dim str_Ret: str_Ret = "1.3" ' default exchange rate

		rs_Prices.Filter = "Period=" & str_Period
		If Not rs_Prices.EOF Then
			str_Ret = rs_Prices.Fields("ExchangeRate")
		End If
		
		GetExchangeRate = str_Ret
	End Function	
	
	' get Transaction Price in USD given the datein period
	Function GetTransactionPrice(str_Period)
		Dim str_Ret: str_Ret = "0"

		rs_Prices.Filter = "Period=" & str_Period
		If Not rs_Prices.EOF Then
			str_Ret = rs_Prices.Fields("TransactionPrice")
		End If
		
		GetTransactionPrice = str_Ret
	End Function		

	' get Aluminum Price in CAD given the aluminum type, supplier, and period	
	Function GetAlumPriceSupplierUS(str_Part, str_Supplier, str_DateIn, str_BundleNos, str_AlumType)
		Dim str_Ret: str_Ret = ""
		Dim a_Date: a_Date = Split(str_DateIn & "", "/")
		Dim str_Period, str_PeriodIn

		Dim a_Bundles: a_Bundles = Split(Replace(str_BundleNos & "",",","/"), "/")
		Dim str_Bundle

		If UBound(a_Bundles) >= 0 Then
			str_Bundle = a_Bundles(0)
		End If

		str_Supplier = Trim(str_Supplier & "")
		
		If UBound(a_Date) >= 2 Then
			str_Period = a_Date(2) & Right("0" & a_Date(0), 2)
		End If

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
					str_Ret = GetPriceUS(str_Supplier, str_Period, str_AlumType)
					str_Debug = str_Debug & ",AT:" & str_AlumType
				Else
					str_AlumType = "_S"
					str_Ret = GetPriceUS(str_Supplier, str_Period, str_AlumType)
					str_Debug = str_Debug & ",AT(Def):" & str_AlumType
				End If
			End If
		End If

		If str_Ret = "" Then
			str_Ret = "0"
			str_Bundles = str_Bundles & str_BundleNos & ","
		End If
		str_Debug = "ID:" & gstr_ID & ",S:" & Show(str_Supplier) & ",&nbsp;CAD$:&nbsp;" & str_Ret & str_Debug & ",P: " + str_Period & ",PO: " + po
		GetAlumPriceSupplierUS = CDbl(str_Ret)
	End Function

	Function Show(str_Val)
		Dim str_Ret: str_Ret = str_Val
		If Trim(str_Val) = "" Then str_Ret = "?"
			Show = str_Ret
	End Function

	Function GetPriceUS(str_Supplier, str_Period, str_AlumType)
		Dim str_Ret: str_Ret = "0"

		rs_Prices.Filter = "Period=" & str_Period
		If Not rs_Prices.EOF Then

			str_Supplier = Replace(str_Supplier & "","-","")

			If str_Supplier & "" <> "" AND str_Supplier <> "KEYMARK" AND str_Supplier <> "SAPA" Then
				str_Ret = rs_Prices.Fields(str_Supplier & str_AlumType)
			Else
				str_Ret = rs_Prices.Fields("Default" & str_AlumType) ' price CAD/kg
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
		
		GetPriceUS = str_Ret
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
	
	Function ConvertToLbFt(str_Kgm, str_Lmm, str_Linch)
		Dim str_Ret,lb,lm,kg,ft
		
		str_Ret = "0"		
		lm = str_Lmm / 1000 ' 1 mm: 0.001 meter
		ft = lm * 3.28084 ' 1 meter : 3.28084 ft
		kg = str_Kgm * lm ' kg value
		lb = kg * 2.20462 ' 1 kg: 2.20462 lb
		
		If lb <> 0 Then
			If str_Linch <> 0 Then 
				str_Ret = lb / (str_Linch / 12) ' use linch to calculate lb/ft
			Else
				str_Ret = lb / ft ' use lmm converted to ft to calculate lb/ft
			End If
		End If
		
		ConvertToLbFt = CDbl(str_Ret)
	End Function
	
	' convert to pounds given the kg
	Function ConvertToLb(str_Kg)
		Dim str_Ret: str_Ret = "0"
		
		str_Ret = str_Kg * 2.20462 ' 1 kg: 2.20462 lb
		
		ConvertToLb = CDbl(str_Ret)
	End Function	
	
	' calculates Transformation Cost in USD
	Function CalculateTransformationCost(str_SupplierPrice, str_Period)
		Dim str_Ret: str_Ret = "0"
		
		' get transaction price in USD
		transactionprice_USD = GetTransactionPrice(str_Period)
			
		' convert transaction price to CAD to get the transformation cost value 
		transactionprice_CAD = transactionprice_USD * GetExchangeRate(str_Period)

		' calculate transformation cost by getting the difference in CAD of transaction price and supplier price
		transformationCost_CAD = transactionprice_CAD - str_SupplierPrice
			
		' convert transformation cost to USD
		transformationCost_USD = transformationCost_CAD / GetExchangeRate(str_Period)			
		str_Ret = transformationCost_USD
		
		CalculateTransformationCost = CDbl(str_Ret)
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