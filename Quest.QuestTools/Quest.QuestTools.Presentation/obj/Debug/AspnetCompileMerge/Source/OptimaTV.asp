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

	sDay = trim(Request.Querystring("sDay"))
	sMonth = trim(Request.Querystring("sMonth"))
	sYear = trim(Request.Querystring("sYear"))
	
if sDay = "" or sMonth = "" or sYear = "" then
	sDay = day(now)
	sMonth = month(now)
	sYear= year(now)

end if
	if Len(sMonth) = 1 then
		sMonth= "0" & sMonth
	end if
	if Len(sDay) = 1 then
		sDay = "0" & sDay
	end if
STAMPVAR = sYear & "-" & sMonth & "-" & sDay




Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Prod; Password=Test123;Initial Catalog=OPTIMA_QUESTWINDOWS"
ON Error goto 0
DBConnection2.Open DSN
%>

</head>
<body>


	  <div class="toolbar">
        <h1 id="pageTitle">Optima</h1>
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
	
	
       
<ul id="screen1" title="Optima" selected="true">
<%
Set rsO = Server.CreateObject("adodb.recordset")
strSQLO = "Select DATACONS, NOTE10, DESCMAT,DESCR_MAT_COMP,DIMENSIONE_X,DIMENSIONE_Y FROM ORDMAST WHERE CONVERT(VARCHAR(10), DATACONS , 23) = '" & STAMPVAR& "' ORDER BY NOTE10 ASC"
rsO.Cursortype = 1
rsO.Locktype = 3
rsO.Open strSQLO, DBConnection2	
'Connect to LASSARD for "Optima Database

OptimizedToday = rsO.RecordCount
%>

		<li class="group">Optimized Today</li>
		<li><% response.write "Total Glass Optimized: " & OptimizedToday %></li>

<%
rsO.filter = ""

	Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
	Response.write "<tr title = 'Click on a Header to Sort by that Column' ><th>Barcode</th><th>Description 1</th><th>Description 2</th><th>Date</th><th>Area</th></tr>"



Do while not rsO.eof

	response.write "<tr>"
		response.write "<td>" & rsO("NOTE10") & "</td>"
		response.write "<td>" & rsO("DESCMAT") & "</td>"
		response.write "<td>" & rsO("DESCR_MAT_COMP") & "</td>"
		response.write "<td>" & rsO("DATACONS") & "</td>"
			GWidth = rsO("DIMENSIONE_X") / 25.4
			GHeight = rsO("DIMENSIONE_Y") / 25.4
			GArea = GWidth * GHeight / 144
		response.write "<td>" & ROund(GArea,2) & "</td>"
		response.write "</tr>"
				



rsO.movenext
loop

rsO.close
set rsO = nothing

%>
</table></li>
<li class="group">Tempered Today</li>
<%
Set rsQ = Server.CreateObject("adodb.recordset")
strSQLQ = "Select Q.VirtMachine, Q.ERRORNUMBER, O.NOTE10, O.DIMENSIONE_X, O.DIMENSIONE_Y FROM QALOG AS Q Inner Join ORDMAST AS O on (O.ID_ORDMAST = Q.ID_ORDMAST) WHERE CONVERT(VARCHAR(10), Q.SERVERDATETIME , 23) = '" & STAMPVAR & "' AND ISNULL(Q.ERRORNUMBER, 0) =0  ORDER BY O.NOTE10 ASC"
rsQ.Cursortype = 1
rsQ.Locktype = 3
rsQ.Open strSQLQ, DBConnection2	
'Connect to LASSARD for "Optima Database

rsQ.filter =" VirtMachine = 'TEMPERING'"
TemperedToday = rsQ.RecordCount
rsQ.filter =""
rsQ.filter =" VirtMachine = 'HS'"
HSToday = rsQ.RecordCount

%>
		<li><% response.write "Total Glass Tempered: " & TemperedToday %></li>
		<li><% response.write "Total Glass HS: " & HSToday %></li>
<%
rsQ.filter = ""

	Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
	Response.write "<tr title = 'Click on a Header to Sort by that Column' ><th>Barcode</th><th>Activity</th><th>Area</th></tr>"



Do while not rsQ.eof

	response.write "<tr>"
		response.write "<td>" & rsQ("NOTE10") & "</td>"
		response.write "<td>" & rsQ("VIRTMachine") & "</td>"
			GWidth = rsQ("DIMENSIONE_X") / 25.4
			GHeight = rsQ("DIMENSIONE_Y") / 25.4
			GArea = GWidth * GHeight / 144
		response.write "<td>" & ROund(GArea,2) & "</td>"
		response.write "</tr>"
				



rsQ.movenext
loop

rsQ.close
set rsQ = nothing



DBConnection.close
set DBConnection=nothing

DBConnection2.Close
set DBConnection2 = nothing
%>


</body>
</html>

