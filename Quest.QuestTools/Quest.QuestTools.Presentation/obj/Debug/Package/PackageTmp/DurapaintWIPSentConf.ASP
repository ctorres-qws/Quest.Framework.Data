<%@ Language=VBScript %>
<html>
<head>
<!-- Changed February 2015 Victor Babuskins / Michael Bernholtz To fix bending line --> 
<!--OLD 8 (Air) 10 (Barcode) 12 (Orderby) -->
<!--NEW 8 (Orderby) 10 (Barcode) 12 (Barcode)   Now-->
<!--NEWER 8 (Orderby) 10 (AIR) 12 (Barcode)-->
<!--#include file="dbpath.asp"-->
<style> TD { mso-number-format:\@; } </style>
</head>
<body>
<%


	DIM str_IDs
	For each item in request.Form("OptimaSelect")
		If str_IDs <> "" Then str_IDs = str_IDs & ","
		str_IDs = str_IDs & item
	Next

	set Rs=server.createobject("ADODB.recordset")
	Rs.open "select * FROM Y_INV WHERE ID IN(" & str_IDs & ") ORDER BY ID ASC",DBConnection,1,3

		Do While Not rs.EOF

			rs.Fields("Note 4") = "SENT"
			Response.write "<br>Marked Sent: "
			Response.write RS("PART") & " " & RS("COLOUR") & " " & RS("PO") & " " & RS("QTY")
			rs.update
			rs.MoveNext
		Loop
		'Next

	DbCloseAll



'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection = nothing
%>
</body>
</html>