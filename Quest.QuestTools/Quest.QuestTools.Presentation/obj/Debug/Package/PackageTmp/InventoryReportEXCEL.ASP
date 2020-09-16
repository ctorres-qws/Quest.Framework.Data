<%@ Language=VBScript %>
<!--#include file="dbpath_secondary.asp"-->
<%

ReportName = Request.Querystring("ReportName")
AlumPrice = Request.Querystring("AlumPrice")

Dim Rs

set Rs=server.createobject("ADODB.recordset")
set Rs2=server.createobject("ADODB.recordset")

Rs.Cursortype = GetDBCursorType
Rs.Locktype = GetDBLockType
Rs.open "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP')  order by PART ASC" ,DBConnection2,1,3
'Rs2.Cursortype = GetDBCursorType
'Rs2.Locktype = GetDBLockType
Rs2.open "SELECT * FROM Y_MASTER order BY ID DESC" ,DBConnection,1,3

'Create a Query
SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
Set RS2 = GetDisconnectedRS(SQL2, DBConnection)
'Get a Record Set

'Aluminum price
alumprice = AlumPrice

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=QCInventoryReport.xls"

If Rs.eof <> true Then

' Introductory Information	
	Response.write "<table>"
	Response.write "<tr><td> Inventory from Goreway / NASHUA / NASHUA PREP / Durapaint / Durapaint (WIP) / Horner </td></tr>"
	Response.write "<tr><td> Using Inventory from: " & ReportName & "</td></tr>"
	Response.write "<tr><td> Aluminium Price: " & AlumPrice & "</td></tr>"

	Response.write "</table>"

' Main Information from Inv SnapShot	
	Response.write "<table border=1>"
	Response.write "<tr><th>Part</th><th>Qty</th><th>Length (mm)</th><th>Colour</th><th>KGM</th><th>Bundle</th><th>Value</th><th>White Paint</th><th>Colour Paint</th></tr>"

	Rs.movefirst

	Do While Not rs.eof
		invpart = trim(rs("part"))
		partqty = rs("Qty")
		lmm = rs("lmm")
		linch = rs("linch")
		colour = rs("colour")
		project = rs("project")
		bundle = rs("bundle")
		kgm = rs("kgm")
		errormsg = 1

		'Rs2.movefirst
		'Do while Not Rs2.eof
		'	If Rs2("Part") = invpart Then
		'		errormsg = 0
		'		If kgm = "0" Then
		'			kgm =0
		'			kgm = rs2("kgm")
		'		End If
		'	End If
		'	Rs2.movenext
		'Loop

	RS2.Filter = "Part='" & invpart & "'"

	if rs2.eof then
		errormsg = 1
	else
		if kgm = "0" then
			kgm =0
			kgm = rs2("kgm")
		end if
			errormsg = 0
	end if
	'rs2.close
	'set rs2 = nothing

		If errormsg = 1 Then
			response.write invpart & " not in inventory master"
		End If

'response.write partqty & "<BR>"
		If Not IsNull(PARTQTY) Then

			If kgm > 5 then 
				pricebar = partqty * kgm
				value2 = value2 + pricebar
				'this code is 5 times overvalued
			Else
				tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
				value = value + tempvalue
			End If

			If colour = "White" Then
				paintlft = (linch/12) * partqty
				totalpaintlft = totalpaintlft + paintlft
			Else

				If colour = "Mill" then
				Else
					paintlft2 = (linch/12) * partqty
					totalpaintlft2 = totalpaintlft2 + paintlft2
				End If

			End If

			Response.write "<tr><td>" & invpart & "</td><td>" & partqty & "</td><td>" & lmm & " </td><td>" & colour & "</td><td>" & kgm & "</td><td>" & bundle & "</td>"

			If kgm > 5 Then
				Response.write "<td>" & round(pricebar,2) & "</td>"
			Else
				Response.write "<td>" & round(tempvalue,2) & "</td>"
			End If
			Response.write "</td>"

			If colour = "White" Then
				Response.write "<td>$" & round(paintlft*0.14,2) & "</td><td></td>"
			Else
				If colour = "Mill" Then
					Response.write "<td>&nbsp;</td><td>&nbsp;</td>"
				Else
					Response.write "<td></td><td>$" & round(paintlft2*0.38,2) & "</td>"
				End If
			End If

			Response.write "</tr>"

		Else
			Response.Write "<td>QTY IN INVENTORY IS ZERO </td>"
		End If
		rs.movenext
	loop

	Response.Write "</table>"

End If

	' Catagorical Conclusions

	paintvalue1 = (totalpaintlft *0.14)
	paintvalue2 = (totalpaintlft2 *0.38)

	Response.write "<table>"
	Response.write "<tr><td>Total kgm Material $" & round(value,2) & "</td></tr>"
	Response.write "<tr><td>Total per bar Material $" & round(value2,2)  & "</td></tr>"
	Response.write "<tr><td>SubTotal = $" & round(value,2) + round(value2,2) & "</td></tr>"
	Response.write "<tr></tr>"
	Response.write "<tr><td>Total Paint White $" & paintvalue1 & "</td></tr>"
	Response.write "<tr><td>Total Paint Project $" & paintvalue2 & "</td></tr>"
	Response.write "<tr><td>Grand Total = $" & round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2) & "</td></tr>"

	Response.write "</table>"

Rs.Close
Set Rs=nothing
Rs2.Close
Set Rs2 = nothing

DBConnection.Close
Set DBConnection=Nothing
DBConnection2.Close
Set DBConnection2=Nothing
%>
