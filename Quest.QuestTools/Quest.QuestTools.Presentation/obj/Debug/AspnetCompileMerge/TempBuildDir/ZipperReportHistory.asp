<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Zipper Reporting Red and Blue - 1 Week worth of Data -->
<!-- Requested by Valdi and Mary Darnell, Michael Bernholtz, May 2017 -->
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Zipper Last 3 Weeks</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Zipper History</h1>
		<a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>

    </div>
<!--#include file="todayandyesterday.asp"-->


<ul id="screen1" title="Last 3 Weeks " selected="true">

	              <li class="group">LAST 3 Weeks</li>
				  <li class="group">Red Zipper</li>
				  
				 <li><table border='1' class='Job' id ='Job' ><thead><tr><th>Week Day</th><th>Date</th><th>Red Zipper</th><th>Frame</th><th>Mullion</th><th>Sill</th><th>Door Sash</th><th>Other</th></tr></thead><tbody>

<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM ProZipperRED ORDER BY PROFILEID ASC"
DaysAgo = 21

Dim dt_Now
dt_Now = Date() - DaysAgo

If (b_SQL_Server) Then
	strSQLSvr = "SELECT '',[Month], [Day], [Year], COUNT(*) as Count_Total, " & vbCrLf
	strSQLSvr = strSQLSvr & "SUM(CASE WHEN Upper(ProfileId) IN('JAMB','ECOWALL FRAME','SLIDE DOOR FRAME','MD FRAME','NC-90 FRAME') THEN 1 ELSE 0 END) Count_Frame, " & vbCrLf
	strSQLSvr = strSQLSvr & "SUM(CASE WHEN Upper(ProfileId) IN('MULLION','ECOWALL MULLION', 'ZIPPER') THEN 1 ELSE 0 END) Count_Mullion, " & vbCrLf
	strSQLSvr = strSQLSvr & "SUM(CASE WHEN Upper(ProfileId) IN('SILL','ECOWALL SILL', 'SILL') THEN 1 ELSE 0 END) Count_Sill, " & vbCrLf
	strSQLSvr = strSQLSvr & "SUM(CASE WHEN Upper(ProfileId) IN('SL DOOR SASH', 'NC 90 SASH', 'MD SASH IN', 'MD SASH OUT') THEN 1 ELSE 0 END) Count_Sash, " & vbCrLf
	strSQLSvr = strSQLSvr & "SUM(CASE WHEN Upper(ProfileId) IN('JAMB','MULLION','SILL','ECOWALL FRAME','SLIDE DOOR FRAME','MD FRAME','NC-90 FRAME','SL DOOR SASH', 'NC 90 SASH', 'ECOWALL MULLION', 'ZIPPER', 'ECOWALL SILL', 'SILL','MD SASH IN', 'MD SASH OUT') THEN 0 ELSE 1 END) Count_Other " & vbCrLf
	strSQLSvr = strSQLSvr & "FROM {TABLE} " & vbCrLf
	strSQLSvr = strSQLSvr & "WHERE [Year] >= " & Year(dt_Now) & " AND [Month] >= " & Month(dt_Now) & vbCrLf
	strSQLSvr = strSQLSvr & "GROUP BY [Month], [Day], [Year] " & vbCrLf
	strSQL = Replace(strSQLSvr,"{TABLE}","ProZipperRed")
End If

'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)

Do Until DaysAgo < 0

	LoopDay = DAY(Date()- DaysAgo)
	LoopWeek = DatePart("ww", Date()- DaysAgo)
	LoopMonth = MONTH(DATE() - DaysAgo)
	LoopYear = YEAR(DATE() - DaysAgo)
	LoopName = Weekday(DATE() - DaysAgo)
	LoopWeekName = WeekDayName(LoopName) 
	Frame = 0
	Mullion = 0
	Sill = 0
	Sash = 0
	Other = 0
	Total = 0
	

	rs.filter = "MONTH = " & loopMonth & " AND YEAR = " & loopyear & " AND DAY = " & LoopDay
	do while not rs.eof
		
	If b_SQL_Server Then
		Frame = RS("Count_Frame")
		Mullion = RS("Count_Mullion")
		Sill = RS("Count_Sill")
		Sash = RS("Count_Sash")
		Other = RS("Count_Other")
		Total = RS("Count_Total")	
	Else
		CurrentProfile = RS("PROFILEID")
		Total = Total + 1
			select Case CurrentProfile
				Case "ecowall frame", "slide door frame", "MD frame", "NC-90 frame"
					Frame = Frame + 1
				Case "ecowall mullion", "zipper"
					Mullion = Mullion + 1
				Case "ecowall sill", "sill"
					Sill = Sill+1
				Case "sl door sash", "nc 90 sash", "MD sash IN", "MD sash OUT" 
					Sash = Sash+1
				Case Else
					Other = Other + 1
			End Select 
	End If
	rs.movenext
	loop

response.write "<tr>"
response.write "<td>" & LoopWeekName & "</td>"
response.write "<td>" & Date()- DaysAgo & "</td>"
response.write "<td><b>" & Total & "</b></td>"
response.write "<td>" & Frame & "</td>"
response.write "<td>" & Mullion & "</td>"
response.write "<td>" & Sill & "</td>"
response.write "<td>" & Sash & "</td>"
response.write "<td>" & Other & "</td>"
response.write "</tr>"

	If DaysAgo > 14 then
		Week1Total = Week1Total + Total
	end if
	
	If DaysAgo > 7 and DaysAgo <= 14 then
		Week2Total = Week2Total + Total
	end if
	
	If DaysAgo <= 7 then
		Week3Total = Week3Total + Total
	end if

DaysAgo = DaysAgo - 1
loop
	
rs.close
set rs=nothing
%>
</table></li>

<li> 7 Day Totals : <%response.write Week1Total%> : <%response.write Week2Total%> : <%response.write Week3Total%> </li>
<li>Grand Total: <%response.write Week1Total + Week2Total + Week3Total %>

<li class="group">Blue Zipper</li>
<li><table border='1' class='Job' id ='Job' ><thead><tr><th>Week Day</th><th>Date</th><th>Blue Zipper</th><th>Jamb</th><th>Mullion</th><th>Sill</th><th>Other</th></tr></thead><tbody>

<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM ProZipperBlue ORDER BY PROFILEID ASC"

If b_SQL_Server Then
	strSQL = 	Replace(strSQLSvr, "{TABLE}", "ProZipperBlue")

End If

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

DaysAgo = 21
Week1Total = 0
Week2Total = 0
Week3Total = 0
Do Until DaysAgo < 0

	LoopDay = DAY(Date()- DaysAgo)
	LoopWeek = DatePart("ww", Date()- DaysAgo)
	LoopMonth = MONTH(DATE() - DaysAgo)
	LoopYear = YEAR(DATE() - DaysAgo)
	LoopName = Weekday(DATE() - DaysAgo)
	LoopWeekName = WeekDayName(LoopName) 
	Frame = 0
	Mullion = 0
	Sill = 0
	Other = 0
	Total = 0
	

	rs.filter = "MONTH = " & loopMonth & " AND YEAR = " & loopyear & " AND DAY = " & LoopDay
	do while not rs.eof
	CurrentProfile = RS("PROFILEID")
	If b_SQL_Server Then
		Frame = rs("Count_Frame")
		Mullion = rs("Count_Mullion")
		Sill = rs("Count_Sill")
		Other = rs("Count_Other")
	Else
	Total = Total + 1
		select Case CurrentProfile
			Case "JAMB"
				Frame = Frame + 1
			Case "MULLION"
				Mullion = Mullion + 1
			Case "SILL"
				Sill = Sill+1
			Case Else
				Other = Other + 1
		End Select 
	End If
	rs.movenext
	loop

response.write "<tr>"
response.write "<td>" & LoopWeekName & "</td>"
response.write "<td>" & Date()- DaysAgo & "</td>"
response.write "<td><b>" & Total & "</b></td>"
response.write "<td>" & Frame & "</td>"
response.write "<td>" & Mullion & "</td>"
response.write "<td>" & Sill & "</td>"
response.write "<td>" & Other & "</td>"
response.write "</tr>"

	If DaysAgo > 14 then
		Week1Total = Week1Total + Total
	end if
	
	If DaysAgo > 7 and DaysAgo <= 14 then
		Week2Total = Week2Total + Total
	end if
	
	If DaysAgo <= 7 then
		Week3Total = Week3Total + Total
	end if	

DaysAgo = DaysAgo - 1
loop

rs.close
set rs=nothing


DBConnection.close
set DBConnection=nothing
%>   
</table></li>
<li> 7 Day Totals : <%response.write Week1Total%> : <%response.write Week2Total%> : <%response.write Week3Total%> </li>
<li>Grand Total: <%response.write Week1Total + Week2Total + Week3Total %>

</ul>
</body>
</html>


