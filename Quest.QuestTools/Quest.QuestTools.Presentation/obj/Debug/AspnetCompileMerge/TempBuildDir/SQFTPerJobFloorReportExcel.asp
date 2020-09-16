<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- June 2016, for Shaun Levy, Created by Michael Bernholtz-->
<!-- Individual Report to describe the SQFT vales per floor of each Job -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Average Openings per Job</title>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_JOBS  ORDER BY JOB ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

    </head>
<body>

<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=SQFT.xls"
%>

        <ul id="Profiles" title="SQFT Per Job/FLoor" selected="true">
<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Total Windows</th><th>Total SQFT</th><th>Average SQFT</th></tr>
<%
Do While Not rs.eof

	Set rsJOB = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM [" & rs("JOB") & "] ORDER BY FLOOR ASC"
	On Error Resume Next  
		rsJOB.Open strSQL, DBConnection
	On Error GoTo 0

	If rsJOB.State = 1 Then 

		WindowNumber = 0
		WindowSQFT = 0
		TotalSQFT = 0
		Floor = ""

		Floor = RSJOB("FLOOR")
		Do while not rsJOB.eof
			oldFloor = FLOOR
			Floor = RSJOB("FLOOR")
			If oldFloor = Floor Then
				WindowNumber = WindowNumber + 1
				WindowSQFT = RSJOB("X") * RSJOB("Y") / 144
				TotalSQFT = TotalSQFT + WindowSQFT
			Else
				Response.write "<tr>"

				If WindowNumber = 0 Then
					AverageSQFT = 0
				Else
					AverageSQFT = TotalSQFT/WindowNumber 
				End If
				Response.write "<td>" & RS("JOB") & "</td><td>" & oldfloor & "</td><td>" & WindowNumber & "</td><td>" & TotalSQFT &"</td><td>" & AverageSQFT & "</td>"
				Response.write "</tr>"

				WindowNumber = 1
				WindowSQFT = RSJOB("X") * RSJOB("Y") / 144
				TotalSQFT = WindowSQFT
			End If

			rsJOB.movenext
		Loop
		Response.write "<tr>"
		Response.write "<td>" & RS("JOB") & "</td><td>" & oldfloor & "</td><td>" & WindowNumber & "</td><td>" & TotalSQFT &"</td><td>" & TotalSQFT/WindowNumber & "</td>"
		Response.write "</tr>"

		rsJOB.Close

	End If
	set rsJOB = nothing
	rs.movenext
Loop
Response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
    </ul>
</body>
</html>
