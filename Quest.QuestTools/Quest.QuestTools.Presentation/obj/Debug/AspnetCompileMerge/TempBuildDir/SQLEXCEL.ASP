<%@ Language=VBScript %>
<%
jobname = request.QueryString("jobname")
fl = request.QueryString("fl")
mode = request.QueryString("mode")
supplier = request.QueryString("supplier")
dim Cn,Rs
set Cn=server.createobject("ADODB.connection")
set Rs=server.createobject("ADODB.recordset")
Cn.open "provider=microsoft.jet.oledb.4.0;data source=" & server.mappath("database2/quest.mdb")

Rs.open "select * FROM Z_GLASSDB ORDER BY ID ASC",Cn,1,3


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=optima.xls"
	if Rs.eof <> true then
	response.write "<table border=1>"

	while not Rs.eof
	
		response.write "<tr><td>" & Rs.fields("ID") & "</td><td>" & Rs.fields("Qty") & "</td><td>" & Rs.fields("1 Mat") & "</td><td>" & Rs.fields("1 Spac") 
			response.write "</td><td>" & Rs.fields("2 Mat") & "</td><td>" & Rs.fields("2 Spac") & "</td><td>" & Rs.fields("3 Mat") & "</td><td>" & Rs.fields("Spacer Distance") 
				response.write "</td><td>" & Rs.fields("UM") & "</td><td>" & Rs.fields("DIM X") & "</td><td>" & Rs.fields("DIM Y") & "</td><td>" & Rs.fields("SINGLE") 
					response.write "</td><td>" & Rs.fields("X1") & "</td><td>" & Rs.fields("Y1") & "</td><td>" & Rs.fields("X2") & "</td><td>" & Rs.fields("Y2") 
						response.write "</td><td>" & Rs.fields("ROT") & "</td><td>" & Rs.fields("PRIOR") & "</td><td>" & Rs.fields("PREF") & "</td><td>" & Rs.fields("RACK") 
							response.write "</td><td>" & Rs.fields("LABELS") & "</td><td>" & Rs.fields("ORDER") & "</td><td>" & Rs.fields("CUSTOMER") & "</td><td>" & Rs.fields("DEL DATE") 
							response.write "</td><td>" & Rs.fields("SPACER TEXT") & "</td><td>" & Rs.fields("SPACER HOLE") & "</td><td>" & Rs.fields("IMPORT") & "</td><td>" & Rs.fields("SAVED") 
							response.write "</td><td>" & Rs.fields("PAR1") & "</td><td>" & Rs.fields("PAR2") & "</td><td>" & Rs.fields("PAR3") & "</td><td>" & Rs.fields("PAR4") 
							response.write "</td><td>" & Rs.fields("PAR5") & "</td><td>" & Rs.fields("PAR6") & "</td><td>" & Rs.fields("PAR7") & "</td><td>" & Rs.fields("PAR8") 
							response.write "</td><td>" & Rs.fields("PAR9") & "</td><td>" & Rs.fields("PAR10") & "</td><td>" & Rs.fields("Rif X1") & "</td><td>" & Rs.fields("Rif Y1") 
							response.write "</td><td>" & Rs.fields("Rif X2") & "</td><td>" & Rs.fields("Rif Y2") & "</td><td>" & Rs.fields("Shape Elaboration") & "</td><td>" & Rs.fields("Note") 
							response.write "</td><td>" & Rs.fields("Note 1") & "</td><td>" & Rs.fields("Note 2") & "</td><td>" & Rs.fields("Note 3") & "</td><td>" & Rs.fields("Note 4") 
							response.write "</td><td>" & Rs.fields("Note 5") & "</td><td>" & Rs.fields("Note 6") & "</td><td>" & Rs.fields("Note 7") & "</td><td>" & Rs.fields("Note 8") 
							response.write "</td><td>" & Rs.fields("Note 9") & "</td><td>" & Rs.fields("Note 10") & "</td><td>" & Rs.fields("Note 11") & "</td><td>" & Rs.fields("Note 12") 
							response.write "</td><td>" & Rs.fields("Note 13") & "</td><td>" & Rs.fields("Note 14") & "</td><td>" & Rs.fields("Note 15") & "</td><td>" & Rs.fields("Note 16") 
							response.write "</td><td>" & Rs.fields("Note 17") & "</td><td>" & Rs.fields("Note 18") & "</td><td>" & Rs.fields("Note 19") & "</td><td>" & Rs.fields("Note 20") 
		response.write "</td></tr>"

	Rs.movenext
	wend
	response.write "</table>"
	end if
	
	
	
	
set rs=nothing
Cn.close
%>
