<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Email at Request of Jody Cash on BaseCamp if this report changes, change StockToday and StockTodayTable July 28th, 2014 -->
<!-- Glass Tools Priority days see note from Jody to build -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<%
	currentDate = Date()

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE SHIPDATE IS NULL ORDER BY [INPUTDATE] DESC"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection
Set rs = GetDisconnectedRS(strSQL, DBConnection)
%>
    </head>
<body>

	  <h2>Service, Commercial, & Production Remake Glass</h2>
        <h3> Stock items as of <% Response.write Date() %> </h3>

<%

'BEGIN OF SASHA PORTION OF REPORT - OPTIMA DATE

rs.filter = "[InputDate] = #" & currentDate & "# AND CompletedDate = NULL AND OptimaDate = NULL"

response.write "<b><u>SASHA - Items added Today - Not Exported to Optima</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if


do while not rs.eof

Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table><br>"

rs.filter = "[OptimaDate] = NULL  AND CompletedDate = NULL"

response.write "<b><u>SASHA - Items Last 3 Days - Not Exported to Optima</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if

do while not rs.eof

if DateDiff("d", rs.fields("InputDate"),currentDate ) > 0  AND DateDiff("d", rs.fields("InputDate"),currentDate ) < 3 then


Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"
end if
rs.movenext
loop
Response.write "</table><br>"


rs.filter = "[OptimaDate] = NULL AND CompletedDate = NULL"


response.write "<b><u>SASHA - Items Last 7 Days - Not Exported to Optima</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if



do while not rs.eof

if  DateDiff("d", rs.fields("InputDate"),currentDate ) >= 3 AND DateDiff("d", rs.fields("InputDate"),currentDate ) < 7 then


Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"




rs.filter = "[OptimaDate] = NULL AND CompletedDate = NULL"

response.write "<b><u>SASHA - Items Older than 7 Days - Not Exported to Optima</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if

do while not rs.eof

if  DateDiff("d", rs.fields("InputDate"),currentDate ) >= 7 then

Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"

' END OF SASHA PORTION OF REPORT   - OPTIMA DATE 
' BEGIN OF KENNY PORTION OF REPORT    - OUTPUT DATE

rs.filter = "[CompletedDate] = NULL "

response.write "<b><u>KENNY - Items Last 10 Days - NOT COMPLETED</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if

do while not rs.eof

if DateDiff("d", rs.fields("OptimaDate"),currentDate ) < 10 then


Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"


rs.filter = "[CompletedDate] = NULL "


response.write "<b><u>KENNY - Items Last 14 Days - NOT COMPLETED</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>PO</th><th>Notes</th></tr>"
end if



do while not rs.eof

if  DateDiff("d", rs.fields("OptimaDate"),currentDate ) >= 10 AND DateDiff("d", rs.fields("OptimaDate"),currentDate ) < 14 then


Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"


rs.filter = "[CompletedDate] = NULL "


response.write "<b><u>KENNY - Items Last 30 Days - NOT COMPLETED</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if



do while not rs.eof

if  DateDiff("d", rs.fields("OptimaDate"),currentDate ) >= 14 AND DateDiff("d", rs.fields("OptimaDate"),currentDate ) < 30 then


Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"




rs.filter = "[CompletedDate] = NULL "

response.write "<b><u>KENNY - Items Older than 30 Days - NOT COMPLETED</u></b>"
RESPONSE.WRITE "<table border='1' class='sortable'>"
if not rs.eof then
	rs.movefirst
	RESPONSE.WRITE "<tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>Type</th><th>Order</th><th>Notes</th></tr>"
end if

do while not rs.eof

if  DateDiff("d", rs.fields("OptimaDate"),currentDate ) >= 30 then

Response.write "<tr>"
Response.write "<td>" & rs("ID") & "</td>"
Response.write "<td>" & rs("Job") & " </td>"
Response.write "<td>" & rs("Floor") & " </td>"
Response.write "<td>" & rs("Tag") & " </td>"
Response.write "<td>" & rs("Dim X") & " </td>"
Response.write "<td>" & rs("Dim Y") & "</td>"
Response.write "<td>" & rs("1 Mat") & " </td>"
Response.write "<td>" & rs("1 Spac") & " </td>"
Response.write "<td>" & rs("2 Mat") & " </td>"
Response.write "<td>" & rs("InputDate") & " </td>"
Response.write "<td>" & rs("OptimaDate") & "</td>"
Response.write "<td>" & rs("RequiredDate") & " </td>"
Response.write "<td>" & rs("CompletedDate") & " </td>"
Response.write "<td>" & rs("Department") & " </td>"
Response.write "<td>" & rs("Orderby") & " </td>"
Response.write "<td>" & rs("Notes") & " </td>"
Response.write "</tr>"

end if

rs.movenext
loop
Response.write "</table><br>"


rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
</body>
</html>
