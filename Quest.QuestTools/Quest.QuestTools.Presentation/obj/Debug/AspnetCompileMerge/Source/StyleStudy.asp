<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->

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
Server.ScriptTimeout=50000
%>
    <%
	
CurrentYear = Year(Now)
CurrentMonth = Month(Now)
TwoAgoYear = Year(Dateadd("M",-2, Now))
TwoAgoMonth = Month(Dateadd("M",-2, Now))
	
	SUCount = 0
	OVCount = 0
	SPCount = 0
	PCount = 0
	SDCount = 0
	SWCount = 0
	SUCount = 0
	MUCount = 0
	
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT JOB,FLOOR,TAG, MONTH, YEAR, DEPT FROM X_BARCODE WHERE YEAR >= " & TwoAgoYear & " AND MONTH>= " & TwoAgoMonth & " AND DEPT = 'GLAZING'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * from Styles"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

do while not rs.eof

	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT  FLOOR, TAG, Style from " & rs("job")
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL, DBConnection
	rs3.filter = "Floor = '" & rs("floor") & "' and Tag = '" & rs("Tag") & "'"

	if rs3.eof then
	else

		rs2.filter = "Name = '" & rs3("style") & "'"
		if rs2.eof then
			response.write "FAIL"
		else
			opening = 1
			do until opening = 9
			Select Case rs2("O"&opening)
				Case "SU"
					SUCount = SUCount + 1
				Case "OV"
					OVCount = OVCount + 1
				Case "SP"
					SPCount = SPCount + 1
				Case "P", "PP"
					PCount = PCount + 1
				Case "SD"
					SDCount = SDCount + 1
				Case "SW"
					SWCount = SWCount + 1
				Case "MU"
					MUCount = MUCount + 1
					if opening = 1 then 
						slot = 0
					else
						slot = opening
					end if
					Select Case rs2("L"&slot)
						Case "SU"
							SUCount = SUCount + 1
						Case "OV"
							OVCount = OVCount + 1
						Case "SP"
							SPCount = SPCount + 1
						Case "P", "PP"
							PCount = PCount + 1
						Case "SD"
							SDCount = SDCount + 1
						Case "SW"
							SWCount = SWCount + 1
					End SELECT
					Select Case rs2("R"&slot)
						Case "SU"
							SUCount = SUCount + 1
						Case "OV"
							OVCount = OVCount + 1
						Case "SP"
							SPCount = SPCount + 1
						Case "P", "PP"
							PCount = PCount + 1
						Case "SD"
							SDCount = SDCount + 1
						Case "SW"
							SWCount = SWCount + 1
					End Select
						
			End Select
			opening = opening +1
			loop
			rs3.filter =""
		end if
		rs2.filter = ""
		
	end if
	
	rs3.close
	set rs3 = nothing
rs.movenext
loop

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Service" selected="true">
        
        
<% 


response.write "<li class='group'>Glass Report Study </li>"
response.write "<li> Number of each openning</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>SU</th><th>OV</th><th>SP</th><th>P or PP</th><th>SW</th><th>SD</th><th>MU</th></tr>"

	response.write "<tr>"
	response.write "<td>" & SUcount & "</td>" 
	response.write "<td>" & OVcount & "</td>" 
	response.write "<td>" & SPCount & "</td>" 
	response.write "<td>" & Pcount & "</td>" 
	response.write "<td>" & SWcount & "</td>" 
	response.write "<td>" & SDcount & "</td>" 
	response.write "<td>" & MUcount & "</td>" 
	response.write " </tr>"

response.write "</table></li></ul>"



rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing
%>
</body>
</html>
