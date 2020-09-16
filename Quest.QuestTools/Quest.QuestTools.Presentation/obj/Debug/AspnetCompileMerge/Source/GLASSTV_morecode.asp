<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodega.asp"-->

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
    SQL = "Select * FROM X_BARCODEGA ORDER BY DATETIME DESC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
	
'Create a Query
    SQL3 = "DELETE * FROM X_BARCODETEMPGA"
'Get a Record Set
    Set RS3 = DBConnection.Execute(SQL3)	
	
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODETEMPGA"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

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
	
	
JOB = REQUEST.QueryString("JOB")
FL = REQUEST.QueryString("FLOOR")

totalg = 0
totala = 0
totalc = 0

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)

Do while not rs.eof
'DATETIME = STAMPVAR
'FILTERDATE = Left(DATETIME, 3)
'This if statement tries to deal with the lenght of the datestamp in characters, 9 and 10 should do it for all year round, including the space or not
'for the latter months or not
DATETIME = rs("DATETIME")

IF month(now) < 10 then
IF Left(DATETIME,9) = STAMPVAR then

IF rs("DEPT") = "GLASSLINE" then
totalg = totalg + 1
end if

end if
  else
  IF Left(DATETIME,10) = STAMPVAR then

IF rs("DEPT") = "GLASSLINE" then
totalg = totalg + 1
end if

end if
  
  
end if  
  rs.movenext
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
		<li><% response.write "GLASSLINE: " & totalg %></li>
	
        
              <li class="group">Latest Activity</li>
        
<%

wcount=0
JFCHECKID=0

rs.movefirst 

do while not rs.eof
rs2.filter = "ID > 0"

		DIM Job, Floor, Dept, JFCHECKID, wcount
		JOB = RS("JOB")
		FLOOR = RS("JOB")
		DEPT = RS("DEPT")
		JFCHECKID = 0
		
		
					do while not rs2.eof
	

					IF rs2("JOB") = RS("JOB") AND rs2("FLOOR") = rs("Floor") AND RS2("DEPT") = rs("DEPT") THEN
					JFCHECKID = RS2("ID")
					wcount = wcount + 1	
					END IF
				rs2.movenext
				LOOP
		
		
					IF JFCHECKID = "0" THEN
					rs2.addnew 
					rs2.fields("JOB") = RS("JOB")
					rs2.fields("FLOOR") = RS("FLOOR")
					rs2.fields("DEPT") = RS("DEPT")
					rs2.fields("TAG") = 1
					RS2.UPDATE
					
					ELSE	
					
					rs2.filter = "ID = " & JFCHECKID
					rs2.fields("TAG") = rs2.Fields("TAG") + wcount
					rs2.update
					
					end if
				
				wcount = 0
				
				
	
	
			
		 %>
<% 
	rs.movenext
loop	
	
	rs2.filter = "ID > 0"
	rs2.movefirst
	
	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
	rs2.movenext
	
	rs4.filter = "ID > 0"	
	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
	'AND ("Floor = '" &  rs2("Floor") & "'")
	'AND "Floor = " rs2("Floor")
	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
	rs2.movenext
	
	rs4.filter = "ID > 0"	
	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
	'AND ("Floor = '" &  rs2("Floor") & "'")
	'AND "Floor = " rs2("Floor")
	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
	rs2.movenext
	
	rs4.filter = "ID > 0"	
	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
	'AND ("Floor = '" &  rs2("Floor") & "'")
	'AND "Floor = " rs2("Floor")
	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
	rs2.movenext
	
	rs4.filter = "ID > 0"	
	rs4.filter = "Job = '" &  rs2("Job") & "' AND Floor = '" &  rs2("Floor") & "'"
	'AND ("Floor = '" &  rs2("Floor") & "'")
	'AND "Floor = " rs2("Floor")
	response.write "<li>" & rs2("DEPT") & " " & rs2("Job") & " " & rs2("Floor") & " " & rs2("Tag") & "/" & rs4("TotalWin") & "</li>"
	rs2.movenext
	
	%> 
        
         <li class="group">Last 5 Scans</li>
        
<%
rs.movefirst 
%>        
        
			<li><% rs5.movefirst
		rs5.filter = "Number = " & rs("employee") 
		IF rs5.bof then
		response.write rs("DEPT") & " " & rs("Employee") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " & rs5("First") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
				end if
				rs.movenext %></li>
                
                	<li><% rs5.filter = "ID > 0"
					rs5.filter = "Number = " & rs("employee") 
					IF rs5.bof then
	response.write rs("DEPT") & " " & rs("Employee") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " & rs5("First") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		end if
				rs.movenext %></li>
                
                                	               	<li><% rs5.filter = "ID > 0"
					rs5.filter = "Number = " & rs("employee") 
					IF rs5.bof then
		response.write rs("DEPT") & " " & rs("Employee") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " & rs5("First") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		end if
				rs.movenext %></li>
                
                                	               	<li><% rs5.filter = "ID > 0"
					rs5.filter = "Number = " & rs("employee") 
					IF rs5.bof then
	response.write rs("DEPT") & " " & rs("Employee") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " & rs5("First") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		end if
				rs.movenext %></li>
                
                                               	<li><% rs5.filter = "ID > 0"
					rs5.filter = "Number = " & rs("employee") 
					IF rs5.bof then
		response.write rs("DEPT") & " " & rs("Employee") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
		else
		response.write rs("DEPT") & " " & rs5("First") & " " & rs("Job") & " " & rs("Floor") & " " & rs("Tag") & " " & rs("Datetime")
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
DBConnection.close
set DBConnection=nothing
%>

