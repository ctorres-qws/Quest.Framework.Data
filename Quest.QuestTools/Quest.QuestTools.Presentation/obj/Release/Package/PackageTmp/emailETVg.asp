<!--#include file="dbpath.asp"-->
     <!--Updated May 2014 to prevent timeout-->                  
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

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select EMPLOYEE, DEPT, FirstComplete FROM X_GLAZING WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " AND (FirstComplete = 'Partial' OR FirstComplete = 'TRUE')"
rs.Cursortype = 1
rs.Locktype = 3
rs.Open strSQL, DBConnection
	

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_EMPLOYEES where Line = 'Glazing'"
rs5.Cursortype = 1
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection
	
totalg = 0
totalg2 = 0

%>
</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Glazing by Employee" selected="true">
    
<!--    <li class="group">Today's Stats</li>-->
    
    <%

DO WHILE NOT RS5.EOF

	TOTALG = 0
	TOTALG2 = 0

	rs.filter = ""
	rs.filter = "EMPLOYEE = '" & RS5("NUMBER") & "' AND FirstComplete = 'TRUE'"
	TotalG = rs.Recordcount

	rs.filter = ""
	rs.filter = "EMPLOYEE = '" & RS5("NUMBER") & "' AND FirstComplete = 'Partial'"
	TotalG2 = rs.Recordcount

Response.write "<LI>" & RS5("Number")
Response.write ": Completed Windows: " & TotalG & " Partial Windows: " & TotalG2 
Response.write "</LI>" 

Response.write "<LI>" 
i = 0
Do until i= TotalG
	if Shift = "0" then 
		response.write "<img src='dkbluesquare.gif'>"  
	else
		response.write "<img src='bluesquare.gif'>"  
	end if
i= i+1 
Loop
i = 0
Do until i= TotalG2
	if Shift = "0" then 
		response.write "<img src='dkredsquare.gif'>"  
	else
		response.write "<img src='redsquare.gif'>"  
	end if
i= i+1 
Loop
Response.write "</LI>" 


' rs.movefirst

	' Do while not rs.eof

		' IF rs("EMPLOYEE") = RS5("NUMBER") then
			' IF UCASE(rs("FirstComplete")) = "TRUE" then
				' totalG = totalG + 1
				' IF RS5("SHIFT") = "0" then
					' 'response.write "<img src='bluesquare.gif'>"
				' ELSE
					' 'response.write "<img src='dkbluesquare.gif'>"
				' END IF
			' ELSE
				' totalG2 = totalG2 + 1
				' IF RS5("SHIFT") = "0" then
					' 'response.write "<img src='redsquare.gif'>"
				' ELSE
					' 'response.write "<img src='dkredsquare.gif'>"
				' END IF
			' END IF
			
		' END IF



	' rs.movenext
	' loop


RS5.MOVENEXT
LOOP


RESPONSE.WRITE "</UL>"

rs.close
set rs=nothing

rs5.close
set rs5=nothing
DBConnection.close
set DBConnection=nothing
%>


</body>
</html>