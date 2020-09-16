<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodeqc.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>


<% 

'Create a Query
    SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
	
'Create a Query
    SQL3 = "DELETE * FROM X_BARCODETEMP1"
'Get a Record Set
    Set RS3 = DBConnection.Execute(SQL3)	
	
	
Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "SELECT * From X_BARCODETEMP1"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * From X_WIN_PROD"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

Set rs5 = Server.CreateObject("adodb.recordset")
strSQL5 = "SELECT * From X_EMPLOYEES"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQL5, DBConnection

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEGA ORDER BY DATETIME DESC"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection
	

JOB = REQUEST.QueryString("JOB")
FL = REQUEST.QueryString("FLOOR")

totalg = 0
totala = 0
totalc = 0
totalsu = 0
totalsp = 0


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not


IF rs("DEPT") = "GLAZING" then
totalg = totalg + 1
end if

IF rs("DEPT") = "ASSEMBLY" then
totala = totala + 1
end if

  rs.movenext
loop


rs6.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs6.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not

IF rs6("DEPT") = "GLASSLINE" AND rs6("TYPE") = "SU" OR rs6("TYPE") = "OV" then
totalsu = totalsu + 1
end if

IF rs6("DEPT") = "GLASSLINE" AND rs6("TYPE") = "SP" then
totalsp = totalsp + 1
end if
 
  rs6.movenext
loop
%>
</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Stats</li>
		<li><% response.write "GLAZING: " & totalg %></li>
		<li><% response.write "ASSEMBLY: " & totala %></li>
        <li><% response.write "IG UNITS: " & totalsu %></li>
        <li><% response.write "SP: " & totalsp %></li>
        <li class="group">Today's Activity</li>
        
<%

wcount=0
JFCHECKID=0


'THIS IS SLOW, BUT IT MAKES SURE THAT ALL WINDOWS RELATED TO A BATCH ARE COUNTED FROM MORE THEN TODAY'S LIST
rs.filter = "ID > 0"
rs.movefirst 

do while not rs.eof
rs7.filter = "ID > 0"


		DIM Job, Floor, Dept, JFCHECKID, wcount
		JOB = RS("JOB")
		FLOOR = RS("JOB")
		DEPT = RS("DEPT")
		JFCHECKID = 0
		
		
					do while not rs7.eof
	

					IF rs7("JOB") = RS("JOB") AND rs7("FLOOR") = rs("Floor") AND rs7("DEPT") = rs("DEPT") THEN
					JFCHECKID = rs7("ID")
					wcount = wcount + 1	
					END IF
				rs7.movenext
				LOOP
		
		
					IF JFCHECKID = "0" THEN
					rs7.addnew 
					rs7.fields("JOB") = RS("JOB")
					rs7.fields("FLOOR") = RS("FLOOR")
					rs7.fields("DEPT") = RS("DEPT")
					rs7.fields("YEAR") = RS("YEAR")
					rs7.fields("MONTH") = RS("MONTH")
					rs7.fields("DAY") = RS("DAY")
					rs7.fields("WEEK") = RS("WEEK")
					rs7.fields("TAG") = 1
					rs7.UPDATE
					
					ELSE	
					
					rs7.filter = "ID = " & JFCHECKID
					rs7.fields("TAG") = rs7.Fields("TAG") + wcount
					rs7.update
					
					end if
				
				wcount = 0
				
				
	
	
			
		 %>
<% 
	rs.movenext
loop	
	rs7.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
	do while not rs7.eof
	
	rs4.filter = "Job = '" &  rs7("Job") & "' AND Floor = '" &  rs7("Floor") & "'"
	if rs4.bof then
	else
	response.write "<li>" & rs7("DEPT") & " " & rs7("Job") & " " & rs7("Floor") & " " & rs7("Tag") & "/" & rs4("TotalWin") & "</li>"
	end if
	rs7.movenext
	
	loop
	
	
	%> 
 
 <li class="group">Last 5 Scans</li>
        
<%
rs.movefirst 
%>        
        
			<li><% rs5.movefirst
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS("EMPLOYEE") %></a><% RESPONSE.Write " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS5("LAST") & "," & RS5("FIRST") %></a><% RESPONSE.Write " " &  rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("TIME")
				end if
				rs.movenext %></li>
					
							<li><% rs5.filter = "ID > 0"
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS("EMPLOYEE") %></a><% RESPONSE.Write " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS5("LAST") & "," & RS5("FIRST") %></a><% RESPONSE.Write " " &  rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("TIME")
				end if
				rs.movenext %></li>
                
                	<li><% rs5.filter = "ID > 0"
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS("EMPLOYEE") %></a><% RESPONSE.Write " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS5("LAST") & "," & RS5("FIRST") %></a><% RESPONSE.Write " " &  rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("TIME")
				end if
				rs.movenext %></li>
                
         	<li><% rs5.filter = "ID > 0"
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS("EMPLOYEE") %></a><% RESPONSE.Write " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS5("LAST") & "," & RS5("FIRST") %></a><% RESPONSE.Write " " &  rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("TIME")
				end if
				rs.movenext %></li>
                
                	<li><% rs5.filter = "ID > 0"
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS("EMPLOYEE") %></a><% RESPONSE.Write " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " %><a href="/FACEBOOK/<% RESPONSE.WRITE RS("EMPLOYEE") %>.JPG" target="_SELF"><% RESPONSE.WRITE RS5("LAST") & "," & RS5("FIRST") %></a><% RESPONSE.Write " " &  rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("TIME")
				end if
				rs.movenext %></li>
                

                
            
        </ul>
        
        
      





</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing
rs4.close
set rs4=nothing
rs5.close
set rs5=nothing
rs6.close
set rs6=nothing
rs7.close
set rs7=nothing

DBConnection.close
set DBConnection=nothing
%>

