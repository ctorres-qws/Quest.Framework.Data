<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Report Designed February 2018 for Jody Cash, Antonio Colalillo, Shaun Levy -->
<!-- Report shows Shipping and Glazing scans for a day -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Discrepency Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="GlazingVShippingDate.asp" target="_self">Date</a>
        </div>

        <ul id="Profiles" title=" Shipping and Glazing" selected="true">
<%
ReportDay = request.querystring("sday")
if ReportDay = "" then
	ReportDay = Day(Date)
end if

ReportMonth = request.querystring("smonth")
if ReportMonth = "" then
	ReportMonth = Month(Date)
end if
ReportYear = request.querystring("syear")
if ReportYear = "" then
	ReportYear = year(Date)
end if


ReportDate = ReportDay & "/" & ReportMonth & "/" & ReportYear
Response.write ReportDate

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT G.JOB, G.FLoor, G.TAG, G.DAY, G.MONTH, G.YEAR, G.DateTime, G.BARCODE, S.BARCODE, S.SHIPDATE FROM X_GLAZING AS G RIGHT JOIN X_SHIPPING AS S ON S.BARCODE = G.Barcode WHERE (G.DAY = " & ReportDay & " and G.Month = " & ReportMonth & " and G.Year = " & ReportYear & ") or (S.ShipDate = #" & ReportDate & "#) ORDER BY S.BARCODE ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

ShipToday = 0
GlazeToday = 0
GlazeOther = 0
GlazeNot = 0


response.write "<li class='group'>Shipping Discrepency Report</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table><tr><TD>"
response.write "<table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Glaze Date</th><th>Shipping Barcode</th><th>Ship Date</th></tr>"
do while not rs.eof
		response.write "<tr><td>" & RS("Job") & "</td><td>" & RS("Floor") & "</td><td>" & RS("Tag") & "</td><td>" & RS("DateTime") & "</td><td>" & RS("BARCODE") & "</td><td>" & RS("SHIPDATE") & "</td>" 
		response.write " </tr>"
		
		ShipToday = ShipToday + 1
		If trim(RS("DAY")) = trim(ReportDay) AND trim(RS("MONTH")) = trim(ReportMonth) AND trim(RS("YEAR")) = trim(ReportYear) then
			GlazeToday = GlazeToday + 1
			
		Else
			If isnull(RS("DAY")) then
				GlazeNot = GlazeNot + 1
			Else
				GlazeOther = GlazeOther + 1
			End if
		End if
		
	rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
</table>
</td> <td valign="top">
<B>
<li>Total Glazing Scans Partial and Full: <%response.write ShipToday%></li>
<li>Glazed Today: <%response.write GlazeToday%></li>
<li>Glazed Previous: <%response.write GlazeOther%></li>
<li>Not Glazed: <%response.write GlazeNot%></li>
<B>
   </td></tr> </table>   
	  </ul>
</body>
</html>
