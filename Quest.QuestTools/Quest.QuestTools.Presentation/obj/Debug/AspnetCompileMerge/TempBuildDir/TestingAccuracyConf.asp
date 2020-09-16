<!--#include file="QCdbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Entry form Confirmation for Testing Accuracy page designed for Daniel Zalcman April 2016-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
INPUTDATE = Request.Querystring("DateIn")
If INPUTDATE = "" THEN
	INPUTDATE = NOW
End if


Ameri20 = REQUEST.QueryString("Ameri20")
Ameri60 = REQUEST.QueryString("Ameri60")
Pro20 = REQUEST.QueryString("Pro20")
Pro60 = REQUEST.QueryString("Pro60")
Pertici20 = REQUEST.QueryString("Pertici20")
Pertici60 = REQUEST.QueryString("Pertici60")
TwoAmeri20 = REQUEST.QueryString("2Ameri20")
TwoAmeri60 = REQUEST.QueryString("2Ameri60")
TwoPertici20 = REQUEST.QueryString("2Pertici20")
TwoPertici60 = REQUEST.QueryString("2Pertici60")
AmeriAdjust = REQUEST.QueryString("AmeriAdjust")
ProAdjust = REQUEST.QueryString("ProAdjust")
PerticiAdjust = REQUEST.QueryString("PerticiAdjust")
TwoAmeriAdjust = REQUEST.QueryString("2AmeriAdjust")
TwoPerticiAdjust = REQUEST.QueryString("2PerticiAdjust")
AmeriMat = REQUEST.QueryString("AmeriMat")
ProMat = REQUEST.QueryString("ProMat")
PerticiMat = REQUEST.QueryString("PerticiMat")
TwoAmeriMat = REQUEST.QueryString("2AmeriMat")
TwoPerticiMat = REQUEST.QueryString("2PerticiMat")

testby = REQUEST.QueryString("testby")
Notes = REQUEST.QueryString("Notes")

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

DBOpenQC DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("INSERT INTO TestingAccuracy ([Ameri20], [Ameri60], [Pro20], [Pro60], [Pertici20], [Pertici60],[2Ameri20], [2Ameri60], [2Pertici20], [2Pertici60], [Date], [AmeriAdjust], [ProAdjust], [PerticiAdjust],[2AmeriAdjust], [2PerticiAdjust], [AmeriMat], [ProMat], [PerticiMat],[2AmeriMat], [2PerticiMat], [testby], [notes]) VALUES( " & Ameri20 & ", " & Ameri60 & ", " & Pro20 & ", " & Pro60 & ", " & Pertici20 & ", " & Pertici60 & ", " & TwoAmeri20 & ", " & TwoAmeri60 & ", " & TwoPertici20 & ", " & TwoPertici60 & ", #" & InputDate & "#, '" & AmeriAdjust & "', '" & ProAdjust & "', '" & PerticiAdjust & "', '" & TwoAmeriAdjust & "', '" & TwoPerticiAdjust & "', '" & AmeriMat & "', '" & ProMat & "', '" & PerticiMat & "', '" & TwoAmeriMat & "', '" & TwoPerticiMat & "', '" & testby & "', '" & notes & "')", isSQLServer)
rs.Cursortype = 2
rs.Locktype = 3
DBConnection.Execute(strSQL)

DbCloseAll

End Function

	
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="TestingAccuracyForm.asp" target="_self">Accuracy</a>
    </div>


    
<ul id="Report" title="Added" selected="true">
<LI> Testing Results for all 5 saws</li>
	<li>Tested by: <%response.write "testby"%> on <%response.write "InputDate"%></li>
	<li>Ameri-can 20: <%response.write ameri20%> </li>
	<li>Ameri-can 60: <%response.write ameri60%> </li>
	<li>Ameri-can Adjustment: <%response.write ameriAdjust%> </li>
	<li>Ameri-can Material: <%response.write ameriMat%> </li>
	
	<li>Proline 20: <%response.write Pro20%> </li>
	<li>Proline 60: <%response.write Pro60%> </li>
	<li>Proline Adjustment: <%response.write ProAdjust%> </li>
	<li>Proline Material: <%response.write ProMat%> </li>
		
		
	<li>Pertici 20: <%response.write Pertici20%> </li>
	<li>Pertici 60: <%response.write Pertici60%> </li>
	<li>Pertici Adjustment: <%response.write PerticiAdjust%> </li>
	<li>Perticit Material: <%response.write PerticitMat%> </li>
	
	
	<li><% response.write "Notes: " & Notes %></li>
	<li><a class="whiteButton" href="TestingAccuracyReport.asp" target="_self">View All Results</a></li>
</ul>


<% 

'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>



