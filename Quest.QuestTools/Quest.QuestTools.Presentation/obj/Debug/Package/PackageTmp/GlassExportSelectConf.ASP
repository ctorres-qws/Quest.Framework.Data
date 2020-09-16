<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<head>
<!-- Changed February 2015 Victor Babuskins / Michael Bernholtz To fix bending line --> 
<!--OLD 8 (Air) 10 (Barcode) 12 (Orderby) -->
<!--NEW 8 (Orderby) 10 (Barcode) 12 (Barcode)   Now-->
<!--NEWER 8 (Orderby) 10 (AIR) 12 (Barcode)-->
<!-- < %@ Language=VBScript %> -->

<style> TD { mso-number-format:\@; } </style>
</head>
<body>
<%
jobname = request.Form("jobname")
fl = request.Form("fl")
mode = request.Form("mode")
supplier = request.Form("supplier")
Dim Rs
Server.ScriptTimeout = 500
Flag = ""
Response.ContentType = "application/vnd.ms-excel"

Response.AddHeader "Content-Disposition", "attachment; filename=SelectedOptima.xls"

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Call Process(false, true)
	Case c_MODE_HYBRID
		Call Process(false,true)
		'Call Process(true,false)
	Case c_MODE_SQL_SERVER
		Call Process(true, true)
End Select

Function Process(isSQLServer, b_Excel)

	DBOpen DBConnection, isSQLServer

	Dim str_IDs
	For each item in request.Form("OptimaSelect")
		If str_IDs <> "" Then str_IDs = str_IDs & ","
		str_IDs = str_IDs & item
	Next

	set Rs=server.createobject("ADODB.recordset")
	
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "select * FROM Z_GLASSDB WHERE ID IN(" & str_IDs & ") ORDER BY ID ASC"
	rs.Cursortype = 1
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	If Rs.eof <> true Then
		If b_Excel Then Response.write "<table border=1>"

		Do While Not rs.EOF

		'For each item in request.Form("OptimaSelect")
			'OptimaSelect = item
			'rs.filter = "ID = " & OptimaSelect

			Flag = Rs.fields("OPTIMADATE")

			' If isNull(Flag) or Flag = "UNSCAN" or Flag = "" Then  
			'Removed September 2019 - Check removed from this page as previous page already does it and it caused problems

' Checks to see if Previously Exported: "OPTIMADATE" is the Optima Date - It will only be filled in after display below
' Added an additional condition for UNSCAN on March 7th 2014

				If b_Excel Then
					Response.write "<tr><td>" & Right(Rs.fields("ID"),3) & "</td><td>" & Rs.fields("Qty") & "</td><td>" & Rs.fields("1 Mat") & "</td><td>" & Rs.fields("1 Spac") 
					Response.write "</td><td>" & Rs.fields("2 Mat") & "</td><td></td><td></td><td>"
					Response.write "</td><td>" & Rs.fields("UM") & "</td><td>" & Rs.fields("DIM X") & "</td><td>" & Rs.fields("DIM Y") & "</td><td>" 
					Response.write "</td><td></td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td>" & Rs.fields("CUSTOMER") & "</td><td>"
					Response.write "</td><td>" & Right(Rs.fields("SPACER TEXT"),3) & "</td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td></td><td>"
					Response.write "</td><td></td><td></td><td></td><td>" & Rs.fields("Note") 
' Commercial Items need Tag and no floor  (Note 2 removed and 3 moved to its place) - Michael Bernholtz September 2014
			TagCheck = Rs.fields("TAG")
			if Left(TagCheck,1) = "-" then
			else
				TagCheck = "-" & TagCheck
			end if

					If Rs.fields("DEPARTMENT") = "Commercial" Then

						Response.write "</td><td>" & Rs.fields("JOB") & "</td><td>" & TagCheck & "</td><td></td><td>" & Rs.fields("SPColour") 
					Else
						Response.write "</td><td>" & Rs.fields("JOB") & "</td><td>" & Rs.fields("FLOOR") & "</td><td>" & TagCheck & "</td><td>" & Rs.fields("SPColour") 
					End If

					Response.write "</td><td>" & Rs.fields("CYCLE") & "</td><td>" & Rs.fields("WINDOW") & "</td><td>" & Rs.fields("NOTES") & "</td><td>" & Rs.fields("ORDERBY") 
					Response.write "</td><td>" & Rs.fields("EMPTY") & "</td><td>" & Rs.fields("BARCODE") & "</td><td>" & Rs.fields("GlassFor") & "</td><td>" & Rs.fields("BARCODE") 
					Response.write "</td><td>" & Rs.fields("PO") & "</td><td>" & Rs.fields("DEPARTMENT") & "</td><td>" & Rs.fields("INPUTDATE") & "</td><td>" & Rs.fields("OPTIMADATE") 
					'Note 17 needs to be blank for Osprey Label
					Response.write "</td><td></td><td>" & Rs.fields("GlassFor") & "</td><td>" & Rs.fields("OverallThickness") & "</td><td>" & Rs.fields("REQUIREDDATE")
				End If

' Added at Request of Gurveen for new Optima additions Feb 2016
' Ext Glass (SpecialStatus / SP )
				ExtStatus = ""

				If inSTR(1, Rs.fields("1 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("1 Mat"),"Q") > 0 or inSTR(1,Rs.fields("1 Mat"),"T") > 0 Then
					ExtStatus = "TMP"
				End If

				If inSTR(1,Rs.fields("1 Mat"),"HS") > 0 Then
					ExtStatus = "HS"
				End If
				ExtSP  = ""

				If Rs.fields("2 Mat") = "" or isNUll(Rs.fields("2 Mat")) Then
					If ExtStatus = "HS" then
						ExtSP = "SP"
					End if
				End if

				If b_Excel Then Response.write "</td><td>" & ExtStatus & "</td><td>" & ExtSP & "</td><td>"
' Spacer (Spandrel Color / Notes - Currently Blank)
				If b_Excel Then Response.write "</td><td></td><td><td>"
' Int Glass (SpecialStatus / SP )
				IntStatus = ""
				If inSTR(1, Rs.fields("2 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("2 Mat"),"Q") > 0 or inSTR(1,Rs.fields("1 Mat"),"T") > 0 then
					IntStatus = "TMP"
				End If

				If inSTR(1,Rs.fields("2 Mat"),"HS") > 0 Then
					IntStatus = "HS"
				End If

				If b_Excel Then Response.write "</td><td>" & IntStatus & "</td><td></td>"
				If b_Excel Then Response.write "<td></td><td></td><td></td>"
				SealedUnit  = ""

				If Rs.fields("1 Mat") <> "" and Rs.fields("2 Mat") <> "" Then
					SealedUnit = "IG"
				End if

				If b_Excel Then Response.write "<td>" & SealedUnit & "</td></tr>"
' Optima Date is now filled with the current Date, and will not appear during the next import as it is no longer null
				rs.Fields("OPTIMADATE") = Date
				rs.update

			'End If
			'removed September 2019
			rs.MoveNext
		Loop
		'Next

		If b_Excel Then Response.write "</table>"
	End If

	DbCloseAll

End Function

'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection = nothing
%>
</body>
</html>