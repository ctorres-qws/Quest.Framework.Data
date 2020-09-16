<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			<!--This is a Stored Procedure that runs every half an hour -->
			<!-- Fills V_Report1 and V_Report2 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
<% Server.ScriptTimeout = 200 %> 
  



<% 
' --------------------------------------------------------------------------------------------------Today
	currentDate = Date
	cDay = Day(currentDate)
	cMonth = Month(currentDate)
	cYear = Year(currentDate )

	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' First set of Todays Data 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	
	
' ---------------------------------------------------------------------------------------------------Collect Glazing and Assembly Data 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE  WHERE (DAY = " & cDAY & " AND MONTH = " & cMONTH & " AND YEAR = " & cYear & " AND DEPT = 'GLAZING') ORDER BY DATETIME DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Do while not rs.eof

Job = RS("JOB")
Floor = RS("FLOOR")
Tag = RS("TAG")

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2= "Select * FROM " & JOB & " where Floor = '" & Floor & "'"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection
rs2.filter = "TAG = '" & Tag & "'"
if rs2.eof then
	OpeningsCount = 8
else
	Style = rs2("Style")
	OpeningsCount = Left(Style,1)
	OpeningsCount = OpeningsCount + 0
end if
rs2.close
set rs2 = nothing

Employee = RS("EMPLOYEE")
Barcode = RS("BARCODE")

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "Select * FROM X_GLAZING"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

rs3.addnew
rs3("Barcode") = RS("Barcode")
rs3("Job") = RS("JOB")
rs3("Floor") = RS("Floor")
rs3("Tag") =  RS("Tag")
rs3("DEPT") = "GLAZING"
rs3("EMPLOYEE") = EMPLOYEE
rs3("Openings") = OpeningsCount
rs3("FirstComplete") = "TRUE"
rs3("Joints") = OpeningsCount * 4
rs3("DateTime") = RS("DATETIME")
rs3("Day") = RS("DAY")
rs3("Month") = RS("MONTH")
rs3("Year") = RS("YEAR")
rs3("Week") = RS("WEEK")
rs3("ONumber") = OpeningsCount
rs3("ScanCount") = 1

x = 1
Do Until x>OpeningsCount
	rs3("O" & x) = EMPLOYEE
X=X+1
loop
rs3.update
rs3.close
set rs3 = nothing

Counter = Counter + 1
  rs.movenext
loop


rs.close
set rs = nothing

DBConnection.close
set DBConnection=nothing



%>
</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle">Collect Data</h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

       <b><u>Today's Activity imported</u></b>
	
		<% response.write Counter %>

	
        </ul>

</body>
</html>
