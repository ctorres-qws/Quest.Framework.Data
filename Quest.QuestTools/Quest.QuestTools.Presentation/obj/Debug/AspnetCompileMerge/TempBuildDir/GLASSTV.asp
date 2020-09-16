<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
	<script src="sorttable.js"></script>
	<!--#include file="dbpath.asp"-->
	<%
	ScanMode = TRUE
	%>
	<!--#include file="Texas_dbpath.asp"-->
  
  <% 

	sDay = trim(Request.Querystring("sDay"))
	sMonth = trim(Request.Querystring("sMonth"))
	sYear = trim(Request.Querystring("sYear"))
	
if sDay = "" or sMonth = "" or sYear = "" then

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
sDay = day(now)
sMonth = month(now)
sYear= year(now)
else

STAMPVAR = sMonth & "/" & sDay & "/" & sYear

end if

Call SetTestDate(sDay, sMonth, sYear)

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEGA WHERE DAY = " & sDAY & " AND MONTH =" & SMonth & " AND YEAR = " & SYear &" ORDER BY DATETIME DESC"

rs.Cursortype = 2
rs.Locktype = 3

if CountryLocation ="USA" then
	rs.Open strSQL, DBConnection_Texas
else
	rs.Open strSQL, DBConnection
end if


'Connect to LASSARD for "Optima Database

Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Prod; Password=Test123;Initial Catalog=OPTIMA_QUESTWINDOWS"
ON Error goto 0
DBConnection2.Open DSN



totalg = 0
totals = 0
totalgf = 0
totalgw = 0
totalgs = 0
totalgf1 = 0
totalgf2 = 0


Do while not rs.eof

	DATETIME = rs("DATETIME")
	GDay = rs("DAY")
	GMonth = rs("MONTH")
	GYear = rs("YEAR")



			IF UCASE(rs("DEPT")) = "FOREL" then
				totalg = totalg + 1
				totalgf = totalgf + 1
			end if

			IF UCASE(rs("DEPT")) = "WILLIAN" then
				totalg = totalg + 1
				totalgw = totalgw + 1
			end if

			IF UCASE(rs("DEPT")) = "SP-SERVICE" then
				totalg = totalg + 1
				totalgs = totalgs + 1
			end if

			IF UCASE(LEFT(rs("BARCODE"),2)) = "GT" then
				totals = totals + 1
			end if
			
			IF UCASE(rs("DEPT")) = "FOREL1" then
				totalg = totalg + 1
				totalgf1 = totalgf1 + 1
			end if
			
			IF UCASE(rs("DEPT")) = "FOREL2" then
				totalg = totalg + 1
				totalgf2 = totalgf2 + 1
			end if


rs.movenext
loop

%>

</head>
<body>


	  <div class="toolbar">
        <h1 id="pageTitle">Glass Produced</h1>
		<% 
		Ticket = Request.QueryString("Ticket") 
		If Ticket = "BarcoderTV" then
			BackButton = "BarcoderTV.asp"
		Else
			if CountryLocation = "USA" then 
				BackButton = "indexTexas.html#_Report"
				HomeSiteSuffix = "-USA"
			else
				BackButton = "index.html#_Report"
				HomeSiteSuffix = ""
			end if 
		End if
				
				
		%>
                <a class="button leftButton" type="cancel" href="<%response.write BackButton%>" target="_self">Reports<%response.write HomeSiteSuffix%></a>
    </div>
	
	
       
<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Today's Production</li>
		<li><% response.write "Total Insulated Glass: " & totalg %></li>
		<% if CountryLocation = "USA" then
		%>
		<li><a href = "GlassTVForel1.asp" target= "_self"> <% response.write "Forel 1 Insulated Glass: " & totalgf1 %></a></li>
		<li><a href = "GlassTVForel2.asp" target= "_self" ><% response.write "Forel 2 Insulated Glass: " & totalgf2 %></a></li>
		<% else 
		%>
		<li><a href = "GlassTVForel.asp" target= "_self"> <% response.write "Forel Insulated Glass: " & totalgf %></a></li>
		<li><a href = "GlassTVWillian.asp" target= "_self" ><% response.write "Willian Insulated Glass: " & totalgw %></a></li>
		<% end if
		%>
		
		<li><a href = "GlassTVServiceBoth.asp" target= "_self" ><% response.write "Service Coded Glass: " & totals %></a></li>
		<li><% response.write "Service Spandrel Glass: " & totalgs %></li>
		
		<li class="group">Today's Scans</li>

<%
rs.filter = ""

	Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
	Response.write "<tr title = 'Click on a Header to Sort by that Column' ><th>Job</th><th>Floor</th><th>Tag</th><th>Opening #</th><th>Type</th><th>PO</th><th>PO Line #</th><th>Department</th><th>TimeStamp</th><th>Barcode</th><th>Optima Width</th><th>Optima Height</th><th>Optima SQFT</th></tr>"



Do while not rs.eof
	DATETIME = rs("DATETIME")
	
	
Set rsO = Server.CreateObject("adodb.recordset")
strSQLO = "Select DIMENSIONE_X, DIMENSIONE_Y, NOTE10 FROM ORDMAST WHERE  NOTE10 ='" & trim(rs("Barcode")) & "' ORDER BY NOTE10 DESC"
rsO.Cursortype = 2
rsO.Locktype = 3
rsO.Open strSQLO, DBConnection2	

'rsO.filter = ""
'rsO.filter = " NOTE10 ='" & trim(rs("Barcode")) & "'"

if not rsO.eof then

	GWidth = rsO("DIMENSIONE_X") / 25.4
	GHeight = rsO("DIMENSIONE_Y") / 25.4
	GArea = GWidth * GHeight / 144


else
	GWidth = "0"
	GHeight = "0"
	GArea = "0"
end if


	
	
	
	IF UCASE(rs("DEPT")) = "FOREL" OR UCASE(rs("DEPT")) = "WILLIAN" OR UCASE(rs("DEPT")) = "SP-SERVICE" OR UCASE(rs("DEPT")) = "FOREL1" OR UCASE(rs("DEPT")) = "FOREL2"  then
		response.write "<tr>"
		response.write "<td>" & rs("JOB") & "</td>"
		response.write "<td>" & rs("FLOOR") & "</td>"
		response.write "<td>" & rs("Tag") & "</td>"
		response.write "<td>" & rs("POSITION") & "</td>"
		response.write "<td>" & rs("Type") & "</td>"
		response.write "<td>" & rs("PO") & "</td>"
		response.write "<td>" & rs("POLINE") & "</td>"
		response.write "<td>" & rs("DEPT") & "</td>"
		response.write "<td>" & rs("DATETIME") & "</td>"
		response.write "<td>" & rs("BARCODE") & "</td>"
		
		BC = rs("BARCODE")
		BarcodeSplit = split(BC,"-")
		CommaCount = Ubound(BarcodeSplit)
		if CommaCount>3 then
		
			response.write "<td>" & BarcodeSplit(4) & "</td>"
			response.write "<td>" & BarcodeSplit(5) & "</td>"
			if isnumeric(BarcodeSplit(4)) = True AND isnumeric(BarcodeSplit(5)) = TRUE then
				response.write "<td>" & BarcodeSplit(4) * BarcodeSplit(5) /144 & "</td>"
			else
				Response.write "<td>Error - Not numbers</td>"
			end if
		else
		
			response.write "<td>" & Round(GWidth,2) & "</td>"
			response.write "<td>" & ROund(GHeight,2) & "</td>"
			response.write "<td>" & ROund(GArea,2) & "</td>"
		
		end if
		
		response.write "</tr>"		
	end if

rsO.close
set rsO = nothing

rs.movenext
loop
Response.write "</table></li>"



rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing

DBConnection_Texas.close
set DBConnection_Texas=nothing

DBConnection2.Close
set DBConnection2 = nothing
%>


</body>
</html>

