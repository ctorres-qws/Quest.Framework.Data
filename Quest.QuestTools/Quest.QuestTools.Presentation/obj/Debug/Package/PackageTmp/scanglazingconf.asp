<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="dbpath.asp"-->
			<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>CDN Glazing</title>
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
EMPLOYEE = request.querystring("EMPLOYEEID")	
bc = request.querystring("window")
Openings = request.querystring("Openings") + 0
jobname = request.querystring("job")
floor = request.querystring("floor")
tag = request.querystring("tag")

ScanCount = 1
OpeningsCount = 0
ScanTime = True
dim OpeningsArray(8)
OpeningsArray(1) = request.querystring("O1")
OpeningsArray(2) = request.querystring("O2")	
OpeningsArray(3) = request.querystring("O3")	
OpeningsArray(4) = request.querystring("O4")	
OpeningsArray(5) = request.querystring("O5")	
OpeningsArray(6) = request.querystring("O6")	
OpeningsArray(7) = request.querystring("O7")	
OpeningsArray(8) = request.querystring("O8")	
dim AllScan(8)
FirstComplete = TRUE
dim Completed(8)
i = 1
Do until i=9
	if Openings >= i then
		Completed(i) = FALSE
	else 
		Completed(i) = TRUE
	end if
i=i+1
loop

i = 1
Do until i = 9
if not OpeningsArray(i) = "" then
OpeningsCount = OpeningsCount + 1
end if
i = i+ 1
loop
	
%>
     <!--#include file="TodayAndYesterday.asp"-->
<%

DEPTVAR = "GLAZING"

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select


Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_Glazing WHERE JOB = '" & jobname & "' and FLOOR = '" & floor & "' and tag = '" & tag & "'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'rs.filter = " JOB = '" & jobname & "' and FLOOR = '" & floor & "' and tag = '" & tag & "'"
if not rs.eof then
	Do until rs.eof
	' NEED TO FINISH THIS for ten minute
'	ccMinute = Right(rs("TIME"),2)
'	ccHour =  Left(rs("TIME"),2)
'	response.write DATEDIFF("M", rs("TIME"), TIME) 
'		if rs("DAY") = DAY(Now) and  rs("MONTH") = MONTH(Now) and  rs("YEAR") = YEAR(Now)  then 
'			ScanTime = False ' Within 10 minutes on the same day
'		end if
		i = 1
		Do until i = 9
			if isnull(Rs("O" & i)) = TRUE or Rs("O" & i) = 0 then
			else
			if Completed(i) = FALSE then
				Completed(i) = TRUE
			end if
				if AllScan(i) = "" then
					AllScan(i) = rs("O" & i)
				else
					AllScan(i) = AllScan(i) & ", " & rs("O" & i)
				end if
			end if
			i = i + 1
		loop
		ScanCount = ScanCount + 1
		if RS("FirstComplete") = "TRUE" then
			FirstComplete = FALSE
		end if
	rs.movenext
	loop
end if
rs.filter = ""


if ScanTime = True then
	if LEN(floor) > 0 then
		if Len(employee) = 4 AND Len(bc) > 5 and openingsCount > 0 then
			rs.addnew 
			rs.fields("BARCODE") = bc
			rs.fields("JOB") = jobname
			rs.fields("FLOOR") = floor
			rs.fields("TAG") = tag
			rs.fields("DEPT") = DEPTVAR
			rs.fields("EMPLOYEE") = EMPLOYEE
			rs.fields("DATETIME") = STAMPVAR
			rs.fields("OPENINGS") = OPENINGS
			rs.fields("JOINTS") = OPENINGS * 4
			rs.fields("TIME") = cctime
			if hour(now) <= 6 then  ' Changed to 6am from 3 by Michael Bernholtz February 2018
				rs.fields("DAY") = cYesterday
				rs.fields("MONTH") = cMonthy
				rs.fields("YEAR") = cYeary
				rs.fields("WEEK") = weekNumbery		
			else
				rs.fields("DAY") = cDay
				rs.fields("MONTH") = cMonth
				rs.fields("YEAR") = cYear
				rs.fields("WEEK") = weekNumber
			end if		
			rs.fields("ONUMBER") = OpeningsCount
			rs.fields("SCANCOUNT") = ScanCount
			x = 1
			Do until x= 9
				rs.fields("O" & x) = OpeningsArray(x)
				if isnull(Rs("O" & x)) = TRUE or Rs("O" & x) = 0 then
				else
					if Completed(x) = FALSE then
						Completed(x) = TRUE
					end if
				end if
			x = x+1
			Loop

			i = 1
			Do until i=9
				if Completed(i) = FALSE then
					FirstComplete = FALSE
				end if
			i=i+1
			loop
	
			if FirstComplete = TRUE then		
				rs.fields("FirstComplete") = "TRUE"
			else
				rs.fields("FirstComplete") = "Partial"
			end if			
			RS.UPDATE

		else 
			error = "Wrong Barcode, Try Again"
		end if ' Employee and Barcode and Openings
	end if ' floor
end if


DbCloseAll

End Function

%>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
     <form id="igline" title="Glazing Scan" class="panel" name="igline" action="ScanGlazing.asp" method="GET" selected="true">

        <fieldset>
       
             <div class="row">
			 <%
			 if Unfound = True then
			 %>
			 <p>Barcode Not found in Database</p>
			 <%
			 end if
			 %>
			 <big>
			 <table border = "1" cellpadding="0" style="solid" align = "center">
			 <tr><th>Opening</th><th>EmpID</th></tr>
			
			 <%
			 i = 1
			 Do until i = Openings + 1
			 
			 
			response.write "<tr>"
			response.write "<td align = 'Center'>" & i & "</td>"
			if Allscan(i) = "" then
			response.write "<td align = 'Center'>" & OpeningsArray(i) & "</td>"
			response.write "<td align = 'Center'>" & Completed(i) & "</td>"
			else
			response.write "<td align = 'Center'>" & AllScan(i) & ",  <b>" & OpeningsArray(i) & "</b></td>"
			response.write "<td align = 'Center'>" & Completed(i) & "</td>"
			end if
			response.write "</tr> "
			 i= i+1
			 loop
			 %>
			</table>
			</big>
<%
if FirstComplete = TRUE then
	response.write "<P> First Complete Scan of " & BC & "</p>"
end if

'rs.close
'set rs = nothing
'DBConnection.close
'set DBConnection = nothing
%>
		<input type="hidden" name="EmployeeID" value="<%response.write employee%>" />
        </fieldset>  
		
		  <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
            </form>
</body>
</html>
