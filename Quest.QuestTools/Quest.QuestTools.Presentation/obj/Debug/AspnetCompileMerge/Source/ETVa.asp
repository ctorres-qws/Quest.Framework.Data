<!--#include file="dbpath.asp"-->
	<!--Cleaned up May 7th-->   

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;

  </script>

<% 

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

Call SetTestDate(cDay, cMonth, cYear)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select EMPLOYEE FROM X_BARCODE WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'ASSEMBLY'"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
	
	
'Create a Query
'    SQL3 = FixSQL("DELETE * FROM X_BARCODETEMPETV")
'Get a Record Set
'    Set RS3 = DBConnection.Execute(SQL3)	
	
	
'Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL2 = "SELECT * From X_BARCODETEMPETV"
'rs2.Cursortype = GetDBCursorType
'rs2.Locktype = GetDBLockType
'rs2.Open strSQL2, DBConnection



Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT xE.* From X_EMPLOYEES xE WHERE Number IN(SELECT Employee FROM x_Barcode WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'ASSEMBLY')"
rs5.Cursortype = GetDBCursorType
rs5.Locktype = GetDBLockType
rs5.Open strSQL5, DBConnection

'Set rs6 = Server.CreateObject("adodb.recordset")
'strSQL6 = "SELECT * From X_BARCODESW"
'rs6.Cursortype = 2
'rs6.Locktype = 3
'rs6.Open strSQL6, DBConnection

'Set rs7 = Server.CreateObject("adodb.recordset")
'strSQL7 = "SELECT * From X_BARCODESD"
'rs7.Cursortype = 2
'rs7.Locktype = 3
'rs7.Open strSQL7, DBConnection

'Set rs8 = Server.CreateObject("adodb.recordset")
'strSQL8 = "SELECT * From X_BARCODEOV"
'rs8.Cursortype = 2
'rs8.Locktype = 3
'rs8.Open strSQL8, DBConnection

TotalA = 0

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Assembly by Employee" selected="true">
    
    <li class="group">Today's Stats: (Light Green = Day Shift / Dark Green = Night Shift)</li>
    
    <%

DO WHILE NOT RS5.EOF

rs.Close
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select COUNT(*) FROM X_BARCODE WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND DEPT = 'ASSEMBLY' AND Employee='" & RS5("NUMBER") & "'"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

'rs.movefirst
TotalA = 0
TotalA = rs(0)

'Do while not rs.eof
'IF rs("EMPLOYEE") = RS5("NUMBER") then
'TotalA = TotalA + 1
'	IF RS5("SHIFT") = "0" then
'	response.write "<img src='greensquare.gif'>"
'	ELSE
'	response.write "<img src='dkgreensquare.gif'>"
'	END IF
'end if
'rs.movenext
'loop

IF TotalA = 0 THEN 
ELSE
	
For i = 0 To TotalA
	IF RS5("SHIFT") = "0" then
	response.write "<img src='greensquare.gif'>"
	ELSE
	response.write "<img src='dkgreensquare.gif'>"
	END IF
Next
	
RESPONSE.WRITE "<li>"
%><a href="ETVadet.asp?ticket=TodayAssembly&employee=<% RESPONSE.WRITE RS5("NUMBER") %>&first=<% response.write rs5("first") %>&last=<% response.write rs5("last") %>" target="_Self"><% RESPONSE.WRITE RS5("NUMBER") & " " & RS5("LAST") & ", " & RS5("FIRST") & " : " & TotalA %></a></li><%
END IF 

RS5.MOVENEXT
LOOP

RESPONSE.WRITE "</UL>"

rs.close
set rs=nothing
'rs2.close
'set rs2=nothing
'rs3.close
'set rs3=nothing
rs5.close
set rs5=nothing
'rs36.close
'set rs6=nothing
'rs7.close
'set rs7=nothing
'rs8.close
'set rs8=nothing
DBConnection.close
set DBConnection=nothing
%>


</body>
</html>