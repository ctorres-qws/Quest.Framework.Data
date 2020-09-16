                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Window Production - Job Floor representation of Window Number and Total SQFT-->
<!-- Collecting from X_WIN_PROD which already stores all of this data-->
<!-- Created July 31st, 2017 by Michael Bernholtz - For Jody Cash and Shaun Levy-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Service Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
Dim str_StartDate, str_EndDate
Dim str_UseTruckCloseDate
str_UseTruckCloseDate = Request("UseTruckClosedDate")


str_StartDate = Request("StartDate")
str_EndDate = Request("EndDate")

If str_StartDate = "" Then str_StartDate = "01/01/" & Year(Now)
If str_EndDate = "" Then str_EndDate = "12/31/" & Year(Now)

Set rs = Server.CreateObject("adodb.recordset")
If str_UseTruckCloseDate = "OK" Then
		strSQL = "SELECT DATESTAMP, JOB, FLOOR, TOTALWIN, TOTALPanels, TOTALDoor, TOTALSQFT FROM X_WIN_PROD WHERE [DATESTAMP]>= #" & str_StartDate & "# AND [DATESTAMP]<= #" & str_EndDate & "# ORDER BY JOB ASC, FLOOR DESC"
Else
	strSQL = strSQL & "SELECT A.DATESTAMP, B.DATESTAMP as DateStamp2,B.JOB, B.FLOOR, B.TOTALWIN, B.TOTALPanels, B.TOTALDoor, B.TOTALSQFT FROM X_WIN_PROD B "
	strSQL = strSQL & "LEFT JOIN (SELECT Job, Floor, Max(ShipDate) as DateStamp FROM x_Shipping_Truck GROUP BY Job, Floor) A ON A.Job = B.Job AND A.Floor = B.Floor "
	strSQL = strSQL & "WHERE [A.DATESTAMP]>= #" & str_StartDate & "# AND [A.DATESTAMP]<= #" & str_EndDate & "#  ORDER BY B.JOB ASC, B.FLOOR DESC "
End If

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
 <style>
 		input { margin-left: 20px !important; width: 300px !important; border: 1px solid rgb(221,221,221) !important; border-radius: 5px; margin-bottom: 5px !important; height: 30px !important; padding-left: 0px !important;}

 	</style>
<script>
	function runReport() {
		document.fMain.submit();
	}
</script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
        <ul id="Profiles" title="Job / Floor - SQFT" selected="true">
       <form name="fMain" method="get">

<%
TotalSQFT = 0
Dim str_Selected
If str_UseTruckCloseDate = "OK" Then str_Selected = " checked=true "
response.write "<li class='group'>JOB / FLOOR - SQFT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li> Start Date(mm/dd/yyyy): <input type=""text"" value='" & str_StartDate & "' name='StartDate'>&nbsp;End Date(mm/dd/yyyy):<input type=""text"" value='" & str_EndDate & "' name='EndDate'><input type='button' name='btnSearch' value='Search' onclick='runReport();'></li>  "
response.write "<li> Use X_Win_Prod Date&nbsp(Default is Truck Closed Date): <input type='checkbox' name='UseTruckClosedDate' value='OK' " & str_Selected & "></li>  "
If str_UseTruckCloseDate = "OK" Then
response.write "<li><table border='1' class='sortable'><tr><th>DATE</th><th>Job</th><th>Floor</th><th>Total Windows</th><th>Total Panels</th><th>Total Doors</th><th>Total SQFT</th></tr>"
Else
	response.write "<li><table border='1' class='sortable'><tr><th>Ship Date</th><th>OE Date</th><th>Job</th><th>Floor</th><th>Total Windows</th><th>Total Panels</th><th>Total Doors</th><th>Total SQFT</th></tr>"
End If
do while not rs.eof
	If str_UseTruckCloseDate = "OK" Then
		response.write "<tr><td>" & RS("DATESTAMP") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TOTALWIN") & "</td><td>" & RS("TOTALPanels") & "</td><td>" & RS("TOTALDoor") & "</td><td>" & RS("TOTALSQFT") & "</td>"
	Else
		response.write "<tr><td>" & RS("DATESTAMP") & "</td><td>" & RS("DATESTAMP2") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TOTALWIN") & "</td><td>" & RS("TOTALPanels") & "</td><td>" & RS("TOTALDoor") & "</td><td>" & RS("TOTALSQFT") & "</td>"
	End If
	response.write " </tr>"
	TotalSQFT = TOTALSQFT + RS("TOTALSQFT")
	rs.movenext
loop
response.write "</table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

response.write "<li>TOTAL SQFT: " & TOTALSQFT & "</li>"
%>
               
<li>//END//</li>
/
</form>
      </ul>
</body>
</html>
