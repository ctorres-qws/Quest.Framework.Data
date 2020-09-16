<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="dbpath.asp"-->
			<!-- Receives info from ScanGlazing -->
			<!-- Feb 2016 Page checks Barocde from Table and from X_Barcode to fill in Employee and Recognize Openings.-->

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
  employeeID = request.QueryString("EmpID")
  if isNumeric(employeeID) then
  else
  employeeID = 0000
  end if
  
  bc = trim(request.QueryString("Window"))
		IsPound = InStr(bc,"#")
		if IsPound>0 then
		bc = left(bc,IsPound-1)
		end if
  jobname = ""
  floor = ""
  tag = ""
  
jobname = Left(bc, 3)
	if inStr(1, bc, "-", 0) = 5 then
		floor = Mid(bc, 4, 1)
		tag = Mid(bc, 5, 8)
	END IF
	if inStr(1, bc, "-", 0) = 6 then
		floor = Mid(bc, 4, 2)
		tag = Mid(bc, 6, 8)
	end if
	if inStr(1, bc, "-", 0) = 7 then
		floor = Mid(bc, 4, 3)
		tag = Mid(bc, 7, 8)
	end if
	if inStr(1, bc, "-", 0) = 8 then
		floor = Mid(bc, 4, 4)
		tag = Mid(bc, 8, 8)
	end if
	if inStr(1, bc, "-", 0) = 9 then
		floor = Mid(bc, 4, 5)
		tag = Mid(bc, 9, 8)
	end if

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select style FROM [" & jobname & "] where FLOOR = '" & floor & "' and TAG = '" & tag & "' order by Floor ASC"
rs.Cursortype = 2
rs.Locktype = 3
On Error Resume Next
rs.Open strSQL, DBConnection
	if Err.Number <> 0 then
		DBConnection.close
		set DBConnection = nothing
		response.write "BARCODE ERROR - PLEASE TYPE IN THIS CODE CORRECTLY"
		response.end
	end if

if rs.eof then
OpenStyle = "8001"
Unfound = True
else
OpeningStyle = rs("style")
OpenStyle = CLng(OpeningStyle)
end if

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM styles where Name = " & OpenStyle & ""
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Open = LEFT(OpenStyle,1)
Dim OpeningEmp(8)

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "Select * FROM X_GLAZING where BARCODE = '" & bc & "' and DEPT = 'GLAZING'"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

if rs3.eof then
	FirstScan = True
	x = 1
	Do until x = 9
	OpeningEmp(x) = 0
	x = x+1
	loop
else
	FirstScan = False
	'rs3.movefirst
	do until rs3.eof
		x = 1
		Do until x = 9
			if isnull(rs3("O" & x)) or rs3("O"& x) = 0 then
			else 
				OpeningEmp(x) = rs3("O" & x)
			end if
			x = x+1
		loop
		rs3.movenext
	loop
end if	
	

   
  %>
  
  
</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ScanGlazing.asp" target="_self">GlazingScan</a>
        </div>
   
   
   
    <form id="igline" title="Glazing Scan" class="panel" name="igline" action="ScanGlazingConf.asp" method="GET" selected="true">

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
			 <tr><th>Opening</th><th>Style</th><th>Now</th><th>Done</th></tr>
			
			 <%
			 response.write "<tr><td align = 'Center' colspan='4' ><b>" & employeeid & "</b></td></tr>"
			 i = 1
			 Do until i = open + 1
			 if OpeningEmp(i) = 0 then
			'response.write "<tr><td align = 'Center'>" & i & "</td><td align = 'Center'>" & rs2("O"&i) & "</td><td align = 'Center'> <input type='text' name='O" & i & "' id='O" & i & "' value = '" & employeeid & "' /></td></tr> "
			response.write "<tr>"
			response.write "<td align = 'Center'>" & i & "</td>"
			response.write "<td align = 'Center'>" & rs2("O"&i) & "</td>"
			response.write "<td align = 'Center'> <input type='checkbox' style='zoom:3' name='O" & i & "' value = '" & employeeID & "' checked /></td>"
			response.write "<td></td>"
			response.write "<td align = 'Center'> <B>" & OpeningEmp(i) & "<B/></td>"
			response.write "</tr> "
			else 
			response.write "<tr><td align = 'Center'>" & i & "</td>"
			response.write "<td align = 'Center'>" & rs2("O"&i) & "</td>"
			response.write "<td align = 'Center'> <input type='checkbox' style='zoom:3' name='O" & i & "' value = '" & employeeID & "' checked /></td>"
			response.write "<td align = 'Center'> <B>" & OpeningEmp(i) & "<B/></td>"
			response.write "</tr> "
			end if
			
			 i= i+1
			 loop
			 
			 %>
			</table>
			</big>
			
<%
rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>	
		<input type="hidden" name="Openings" value="<%response.write open%>" />
		<input type="hidden" name="Window" value="<%response.write bc%>" />
		<input type="hidden" name="Job" value="<%response.write jobname%>" />
		<input type="hidden" name="Floor" value="<%response.write floor%>" />
		<input type="hidden" name="Tag" value="<%response.write tag%>" />
		<input type="hidden" name="EmployeeID" value="<%response.write employeeID%>" />
        </fieldset>
        <BR>
		
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
         
            </form>
	

</body>
</html>