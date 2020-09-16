<%@ Language=VBScript %>
<!-- Changed February 2015 Victor Babuskins / Michael Bernholtz To fix bending line --> 
<!--OLD 8 (Air) 10 (Barcode) 12 (Orderby) -->
<!--NEW 8 (Orderby) 10 (Barcode) 12 (Barcode)   Now-->
<!--NEWER 8 (Orderby) 10 (AIR) 12 (Barcode)-->
<!--#include file="dbpath.asp"-->
<%
jobname = request.QueryString("jobname")
fl = request.QueryString("fl")
mode = request.QueryString("mode")
supplier = request.QueryString("supplier")
dim Rs
set Rs=server.createobject("ADODB.recordset")

Rs.open "select * FROM Z_GLASSDB ORDER BY ID ASC",DBConnection,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=optima.xls"
If Rs.eof <> true Then
		response.write "<table border=1>"

	While Not Rs.eof 
	Flag = Rs.fields("OPTIMADATE")
		If isNull(Flag) or Flag = "UNSCAN" Then
' Checks to see if Previously Exported: "OPTIMADATE" is the Optima Date - It will only be filled in after display below
' Added an additional condition for UNSCAN on March 7th 2014	
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
			If Rs.fields("DEPARTMENT") = "Commercial" Then
				Response.write "</td><td>" & Rs.fields("JOB") & "</td><td>" & Rs.fields("TAG") & "</td><td></td><td>" & Rs.fields("CONDITION") 
			Else

			Response.write "</td><td>" & Rs.fields("JOB") & "</td><td>" & Rs.fields("FLOOR") & "</td><td>" & Rs.fields("TAG") & "</td><td>" & Rs.fields("CONDITION") 
		end if
		Response.write "</td><td>" & Rs.fields("CYCLE") & "</td><td>" & Rs.fields("WINDOW") & "</td><td>" & Rs.fields("NOTES") & "</td><td>" & Rs.fields("ORDERBY") 
		Response.write "</td><td>" & Rs.fields("EMPTY") & "</td><td>" & Rs.fields("BARCODE") & "</td><td>" & Rs.fields("WindowNumber") & "</td><td>" & Rs.fields("BARCODE") 
		Response.write "</td><td>" & Rs.fields("PO") & "</td><td>" & Rs.fields("DEPARTMENT") & "</td><td>" & Rs.fields("INPUTDATE") & "</td><td>" & Rs.fields("OPTIMADATE") 
		Response.write "</td><td>" & Rs.fields("REQUIREDDATE") & "</td><td>" & Rs.fields("COMPLETEDDATE") & "</td><td>" & Rs.fields("RECEIVEDDATE") & "</td><td>" & Rs.fields("HIDE") 
		' Added at Request of Gurveen for new Optima additions Feb 2016
		' Ext Glass (SpecialStatus / SP )
		ExtStatus = ""

		If inSTR(1, Rs.fields("1 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("1 Mat"),"Q") > 0 Then
			ExtStatus = "TMP"
		End If

		If inSTR(1,Rs.fields("1 Mat"),"HS") > 0 Then
			ExtStatus = "HS"
		End If
		ExtSP  = ""

		If Rs.fields("2 Mat") = "" or isNUll(Rs.fields("2 Mat")) Then
			ExtSP = "SP"
		End if

		Response.write "</td><td>" & ExtStatus & "</td><td>" & ExtSP & "</td><td>"
		' Spacer (Spandrel Color / Notes - Currently Blank)
		Response.write "</td><td></td><td><td>"
		' Int Glass (SpecialStatus / SP )
		IntStatus = ""

		If inSTR(1, Rs.fields("2 Mat"),"TMP") > 0 or inSTR(1,Rs.fields("2 Mat"),"Q") > 0 Then
			IntStatus = "TMP"
		End If

		If inSTR(1,Rs.fields("2 Mat"),"HS") > 0 Then
			IntStatus = "HS"
		End If

		Response.write "</td><td>" & IntStatus & "</td><td></td>"
		Response.write "<td></td><td></td><td></td>"
		SealedUnit  = ""

		If Rs.fields("1 Mat") <> "" and Rs.fields("2 Mat") <> "" Then
			SealedUnit = "IG"
		End if
		Response.write "<td>" & SealedUnit & "</td></tr>"		

'rs.Fields("OPTIMADATE") = Date
' Optima Date commented out for SAFE viewing.
' This Page will allow an Excel document to be viewed without sending to OPTIMA
'Testing purposes

		End If
		Rs.movenext
	Wend
	Response.write "</table>"
End If

set rs=nothing
DBConnection.close
Set DBConnection = Nothing
%>
