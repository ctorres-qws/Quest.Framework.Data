<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Panel information Cloned from Glass LIne -->
<!-- Michael Bernholtz, January 2015 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  <script src="sorttable.js"></script>
  

<% 

	sDay = trim(Request.Querystring("sDay"))
	sMonth = trim(Request.Querystring("sMonth"))
	sYear = trim(Request.Querystring("sYear"))

if sDay = "" or sMonth = "" or sYear = "" then
sYear = year(now)
sMonth = month(now)
sDay = day(now)
STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)

else

STAMPVAR = sMonth & "/" & sDay & "/" & sYear

end if

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEP WHERE DEPT IN ('Cut','Bend','Ship','Receive') AND [Year]=" & sYear & " AND [Month]=" & sMonth & " AND [Day]=" & sDay & " ORDER BY JOB ASC, Floor ASC, DATETIME DESC"
'rs.Cursortype = GetDBCursorType
'rs.Locktype = GetDBLockType
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

totalp = 0

Do while not rs.eof

	DATETIME = rs("DATETIME")

	' Changed from 9 or 10, because single day month and single day year can make 9-8 or 10-9
		IF STAMPVAR = Left(DATETIME,8) OR STAMPVAR = Left(DATETIME,9) OR STAMPVAR = Left(DATETIME,10) then
				totalp = totalp + 1
		end if
rs.movenext
loop
'rs.movefirst
%>

</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
		<% 
			Ticket = Request.QueryString("Ticket") 
			If Ticket = "BarcoderTV" then
			BackButton = "BarcoderTV.asp"
			Else
			BackButton = "index.html#_Report"
			End if
		%>
                <a class="button leftButton" type="cancel" href="<%Response.Write BackButton %>" target="_self">Reports</a>
				<a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Today's Production</li>
		<li><% response.write "Total Panels: " & totalp %></li>
		<li class="group">Today's Scans</li>

<%


	Response.write "<li><table border ='1' class='sortable' cellpadding='3'>"

	if not rs.eof then	
rs.movefirst
end if
Cut = 0
CM1 =0 ' EuroMac Alum
CM2 = 0 ' EuroMac Steel
CM3 = 0 ' Durma
Bend = 0
BM1 = 0 ' All Steel
BM2 = 0 ' Toskar
BM3= 0 ' Schechtl
Ship = 0
Receive = 0	

Do while not rs.eof
	DATETIME = rs("DATETIME")
	
		IF Left(DATETIME,8) = STAMPVAR OR Left(DATETIME,9) = STAMPVAR OR Left(DATETIME,10) = STAMPVAR then

			
			Select Case rs("DEPT")
				Case "Cut"
					Cut = Cut + 1
					Select Case RS("EMPLOYEE")
						Case "EuroMac - Alum"
							CM1 = CM1 + 1
						Case "EuroMac - Steel"
							CM2 = CM2 + 1
						Case "Durma"
							CM3 = CM3 + 1
					End Select
					
				Case "Bend"
					Bend = Bend + 1	
					Select Case RS("EMPLOYEE")
						Case "All Steel"
							BM1 = BM1 + 1
						Case "Toskar"
							BM2 = BM2 + 1
						Case "Schechtl"
							BM3 = BM3 + 1
					End Select
				Case "Ship"
					Ship = Ship + 1
				Case "Receive"
					Receive = Receive + 1
			End Select
		end if

rs.movenext
loop
'rs.close
'set rs=nothing
response.write "<li><table border='1' class = 'sortable' ><tr><th>Job / FLoor</th><th>Cut</th><th>Euro Alum</th><th>Euro Steel</th><th>Durma</th><th>Bend</th><th>All Steel</th><th>Toskar</th><th>Schechtl</th><th>Ship</th><th>Receive</th></tr>"

'strSQL = "Select * FROM X_BARCODEP WHERE DAY = " & DAY(NOW) & " AND MONTH = " & MONTH(NOW) & " AND YEAR = " & YEAR(NOW) & " ORDER BY JOB ASC, FLOOR ASC, DEPT ASC, TAG ASC"
'Set rs = Server.CreateObject("adodb.recordset")
'rs.Cursortype = GetDBCursorType
'rs.Locktype = GetDBLockType
'rs.Open strSQL, DBConnection
rs.movefirst

Job1 = ""
Job2 = ""
Cut = 0
CM1 = 0
CM2 = 0
CM3 = 0
Bend = 0
BM1 = 0
BM2 = 0
BM3 = 0
Ship = 0
Receive = 0	
TotalCut = 0
TotalCM1 = 0
TotalCM2 = 0
TotalCM3 = 0
TotalBend = 0
TotalBM1 = 0
TotalBM2 = 0
TotalBM3 = 0
TotalShip = 0
TotalReceive = 0
If Not rs.EOF Then
Job1 = rs("JOB") & RS("FLOOR")
Job2 = rs("JOB") & RS("FLOOR")
do while not rs.eof
	Job2 = Job1
	Job1 = rs("JOB") & RS("FLOOR")
	Select Case rs("DEPT")
			Case "Cut"
				Cut = Cut + 1
				TotalCut = TotalCut + 1
				Select Case RS("EMPLOYEE")
						Case "EuroMac - Alum"
							CM1 = CM1 + 1
							TotalCM1 = TotalCM1 + 1
						Case "EuroMac - Steel"
							CM2 = CM2 + 1
							TotalCM2 = TotalCM2 + 1
						Case "Durma"
							CM3 = CM3 + 1
							TotalCM3 = TotalCM3 + 1
					End Select
			Case "Bend"
				Bend = Bend + 1
				TotalBend = TotalBend + 1
					Select Case RS("EMPLOYEE")
						Case "All Steel"
							BM1 = BM1 + 1
							TotalBM1 = TotalBM1 + 1
						Case "Toskar"
							BM2 = BM2 + 1
							TotalBM2 = TotalBM2 + 1
						Case "Schechtl"
							BM3 = BM3 + 1
							TotalBM3 = TotalBM3 + 1
					End Select
			Case "Ship"
				Ship = Ship + 1
				TotalShip = TotalShip + 1
			Case "Receive"
				Receive = Receive + 1
				TotalReceive = TotalReceive + 1
		End Select
	if Job1 = Job2 then	
	else
	response.write "<tr>"
	response.write "<td><b>" & Job2 & "</b></td>"
	response.write "<td><b>" & cut & "</b></td>"
	response.write "<td>" & CM1 & "</td>"
	response.write "<td>" & CM2 & "</td>"
	response.write "<td>" & CM3 & "</td>"
	response.write "<td><b>" & bend & "</b></td>"
	response.write "<td>" & BM1 & "</td>"
	response.write "<td>" & BM2 & "</td>"
	response.write "<td>" & BM3 & "</td>"
	response.write "<td><b>" & ship & "</b></td>"
	response.write "<td><b>" & receive & "</b></td>"
	response.write "</tr>"
	
	Cut = 0
	CM1 = 0
	CM2 = 0
	CM3 = 0
	Bend = 0
	BM1 = 0
	BM2 = 0
	BM3 = 0
	Ship = 0
	Receive = 0	
	end if
	
	
rs.movenext
loop
	response.write "<tr>"
	response.write "<td><b>" & Job2 & "</b></td>"
	response.write "<td><b>" & cut & "</b></td>"
	response.write "<td>" & CM1 & "</td>"
	response.write "<td>" & CM2 & "</td>"
	response.write "<td>" & CM3 & "</td>"
	response.write "<td><b>" & bend & "</b></td>"
	response.write "<td>" & BM1 & "</td>"
	response.write "<td>" & BM2 & "</td>"
	response.write "<td>" & BM3 & "</td>"
	response.write "<td><b>" & ship & "</b></td>"
	response.write "<td><b>" & receive & "</b></td>"
	response.write "</tr>"
End If	

response.write "</table></li>"



DBConnection.close
set DBConnection=nothing
%>
<li>
<table border ='1' class='sortable' cellpadding='3'>
<tr><th>Totals</th></tr>
<tr><th>Cut</th><th>Euro Alum</th><th>Euro Steel</th><th>Durma</th><th>Bend</th><th>All Steel</th><th>Toskar</th><th>Schechtl</th><th>Ship</th><th>Receive</th></tr>
<tr><td><%response.write Totalcut%></td><td><%response.write TotalCM1%></td><td><%response.write TotalCM2%></td><td><%response.write TotalCM3%></td>
<td><%response.write TotalBend%></td><td><%response.write TotalBM1%></td><td><%response.write TotalBM2%></td><td><%response.write TotalBM3%></td>

<td><%response.write TotalShip%></td><td><%response.write TotalReceive%></td></tr>
</table></li>
</ul>

</body>
</html>

