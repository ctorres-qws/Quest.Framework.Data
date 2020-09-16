<%@ Language=VBScript %>
<html>
<head>
<!-- Changed February 2015 Victor Babuskins / Michael Bernholtz To fix bending line --> 
<!--OLD 8 (Air) 10 (Barcode) 12 (Orderby) -->
<!--NEW 8 (Orderby) 10 (Barcode) 12 (Barcode)   Now-->
<!--NEWER 8 (Orderby) 10 (AIR) 12 (Barcode)-->

<style> TD { mso-number-format:\@; } </style>
</head>
<body>
<%
jobname = request.QueryString("jobname")
fl = request.QueryString("fl")
mode = request.QueryString("mode")
supplier = request.QueryString("supplier")
dim Cn,Rs
set Cn=server.createobject("ADODB.connection")
set Rs=server.createobject("ADODB.recordset")
Cn.open "provider=microsoft.jet.oledb.4.0;data source=" & server.mappath("database2/quest.mdb")

Rs.open "select * FROM Z_GLASSDB WHERE [DEPARTMENT] = 'Commercial' ORDER BY ID ASC",Cn,1,3


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=CommercialOptima.xls"
	if Rs.eof <> true then
	response.write "<table border=1>"

	while not Rs.eof 
	Flag = Rs.fields("OPTIMADATE")
		if isNull(Flag) or Flag = "UNSCAN" then
' Checks to see if Previously Exported: "OPTIMADATE" is the Optima Date - It will only be filled in after display below
' Added an additional condition for UNSCAN on March 7th 2014	
		response.write "<tr><td>" & Right(Rs.fields("ID"),3) & "</td><td>" & Rs.fields("Qty") & "</td><td>" & Rs.fields("1 Mat") & "</td><td>" & Rs.fields("1 Spac") 
				response.write "</td><td>" & Rs.fields("2 Mat") & "</td><td></td><td></td><td>"
					response.write "</td><td>" & Rs.fields("UM") & "</td><td>" & Rs.fields("DIM X") & "</td><td>" & Rs.fields("DIM Y") & "</td><td>" 
						response.write "</td><td></td><td></td><td></td><td>"
							response.write "</td><td></td><td></td><td></td><td>"
								response.write "</td><td></td><td></td><td>" & Rs.fields("CUSTOMER") & "</td><td>"
								response.write "</td><td>" & Right(Rs.fields("SPACER TEXT"),3) & "</td><td></td><td></td><td>"
								response.write "</td><td></td><td></td><td></td><td>"
								response.write "</td><td></td><td></td><td></td><td>"
								response.write "</td><td></td><td></td><td></td><td>"
								response.write "</td><td></td><td></td><td></td><td>" & Rs.fields("Note") 
								response.write "</td><td>" & Rs.fields("JOB") & "</td><td>" & Rs.fields("FLOOR") & "</td><td>" & Rs.fields("TAG") & "</td><td>" & Rs.fields("SPColour") 
								response.write "</td><td>" & Rs.fields("CYCLE") & "</td><td>" & Rs.fields("WINDOW") & "</td><td>" & Rs.fields("NOTES") & "</td><td>" & Rs.fields("ORDERBY") 
								response.write "</td><td>" & Rs.fields("EMPTY") & "</td><td>" & Rs.fields("BARCODE") & "</td><td>" & Rs.fields("GlassFor") & "</td><td>" & Rs.fields("BARCODE") 
								response.write "</td><td>" & Rs.fields("PO") & "</td><td>" & Rs.fields("DEPARTMENT") & "</td><td>" & Rs.fields("INPUTDATE") & "</td><td>" & Rs.fields("OPTIMADATE") 
								response.write "</td><td>" & Rs.fields("REQUIREDDATE") & "</td><td>" & Rs.fields("GlassFor") & "</td><td>" & Rs.fields("OverallThickness") & "</td><td>" & Rs.fields("HIDE") 
											' Added at Request of Gurveen for new Optima additions Feb 2016
								' Ext Glass (SpecialStatus / SP )
									ExtStatus = ""
									if inSTR(1, Rs.fields("1 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("1 Mat"),"Q") > 0 or inSTR(1,Rs.fields("1 Mat"),"T") > 0 then 
										ExtStatus = "TMP"
									end if
									if inSTR(1,Rs.fields("1 Mat"),"HS") > 0 then 
										ExtStatus = "HS"
									end if
									ExtSP  = ""
									if Rs.fields("2 Mat") = "" or isNUll(Rs.fields("2 Mat")) then
										IF ExtStatus = "HS" then
											ExtSP = "SP"
										End if
									End if
								response.write "</td><td>" & ExtStatus & "</td><td>" & ExtSP & "</td><td>"
								' Spacer (Spandrel Color / Notes - Currently Blank)
								response.write "</td><td></td><td><td>"
								' Int Glass (SpecialStatus / SP )
									IntStatus = ""
									if inSTR(1, Rs.fields("2 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("2 Mat"),"Q") > 0 or inSTR(1,Rs.fields("1 Mat"),"T") > 0  then 
										IntStatus = "TMP"
									end if
									if inSTR(1,Rs.fields("2 Mat"),"HS") > 0 then 
										IntStatus = "HS"
									end if
								response.write "</td><td>" & IntStatus & "</td><td></td>"
								response.write "<td></td><td></td><td></td>"
								SealedUnit  = ""
									if Rs.fields("1 Mat") <> "" and Rs.fields("2 Mat") <> "" then
										SealedUnit = "IG"
									End if
								response.write "<td>" & SealedUnit & "</td></tr>"	
		rs.Fields("OPTIMADATE") = Date
' Optima Date is now filled with the current Date, and will not appear during the next import as it is no longer null
		end if
	Rs.movenext
	wend
	response.write "</table>"
	end if
	
	
	
rs.close	
set rs=nothing
Cn.close
set Cn = nothing
%>
</body>
</html>