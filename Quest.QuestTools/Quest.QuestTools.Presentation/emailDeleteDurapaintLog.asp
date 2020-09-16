<!--#include file="dbpath.asp"-->
     <!--Updated May 2014 to prevent timeout-->                  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 

  <% 

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM Y_INVLOG WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND Warehouse = 'DURAPAINT' AND Transaction = 'exit' ORDER BY PART"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
	</head>
<body>


<ul id="screen1" title="Deleted Today from Durapaint" selected="true">
    
<li>Deleted Today from Durapaint in total</li>
    
 <%

if rs.eof then
	Response.write "<li>No items Deleted Today from Durapaint</li>"
else
	Part1 = rs("part")
	Part2 = "0"
	PartCount = 0
	Do while not rs.eof
		Part2 = Part1
		Part1 = rs("part")
		if Part2 = Part1 then
			PartCount = PartCount + 1
		else
			response.write "<li>" & UCASE(Part2) & ": " & PartCount & "</li>"
			PartCount = 1
		end if
	rs.movenext
	loop
		response.write "<li>" & UCASE(Part1) & ": " & PartCount & "</li>"
end if

RESPONSE.WRITE "</UL>"

rs.close
set rs=nothing
%>

<br>
<ul id="screen1" title="Added New to Goreway Today" selected="true">
<li>Added New to Goreway Today</li>
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM Y_INVLOG WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND Warehouse = 'GOREWAY' AND Transaction = 'enter' ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

if rs2.eof then
	Response.write "<li>No items added to Goreway</li>"
else
	Part1 = rs2("part")
	Part2 = "0"
	PartCount = 0
	Do while not rs2.eof
		Part2 = Part1
		Part1 = rs2("part")
		if Part2 = Part1 then
			PartCount = PartCount + 1
		else
			response.write "<li>" & UCASE(Part2) & ": " & PartCount & "</li>"
			PartCount = 1
		end if
	rs2.movenext
	loop
		response.write "<li>" & UCASE(Part1) & ": " & PartCount & "</li>"
end if

RESPONSE.WRITE "</UL>"

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>