                      
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->

<!-- Testing Results stored in the system - Designed for Daniel Zalcman - April 2016, Michael Bernholtz-->
<!-- Main Page -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Accuracy Test</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

    <%
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM TestingAccuracy ORDER BY DATE DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_QCT" target="_self">QC Tests</a>
        </div>
      
       
        <ul id="Profiles" title="Machine Accuracy " selected="true">
        <li class='group'>Accuracy Tests</li>
		<a class="whiteButton" href="TestingAccuracyForm.asp" target='_Self'>Add New Result</a>
<%

response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
'response.write "<li><table border='1' class='sortable'  style=' width: 100%'><thead><tr><th>Date</th><th>Ameri-can 20</th><th>Ameri-Can 60</th><th>Material</th><th>Adjustment</th><th>Proline 20</th><th>Proline 60</th><th>Material</th><th>Adjustment</th><th>Pertici 20</th><th>Pertici 60</th><th>Material</th><th>Adjustment</th><th>Ameri-2 20</th><th>Ameri-2 60</th><th>Material</th><th>Adjustment</th><th>Pert-2 20</th><th>Pert-2 60</th><th>Material</th><th>Adjustment</th><th>Tested By</th><th>Notes</th></tr></thead><tbody>"
response.write "<li><table border='1' class='sortable'  style=' width: 100%'><thead><tr><th>Date</th><th>Ameri 20</th><th>Ameri 60</th><th>Adj</th><th>Proline 20</th><th>Proline 60</th><th>Adj</th><th>Pert 20</th><th>Pert 60</th><th>Adj</th><th>Ameri-2 20</th><th>Ameri-2 60</th><th>Adj</th><th>Pert-2 20</th><th>Pert-2 60</th><th>Adj</th><th>Tested By</th><th>Notes</th></tr></thead><tbody>"

'Added Commercial Saws, and needed to reduce table by shortening names and removing Material from this report. It is still collected in DB, April 2017


if rs.eof then
Response.write "<tr><td colspan ='10'>No current Tests</td></tr>"
end if	
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Date") & "</td>"
	response.write "<td>" & RS("Ameri20") &"</td>"
	response.write "<td>" & RS("Ameri60") &"</td>"
	'response.write "<td>" & RS("AmeriMat") &"</td>"
	response.write "<td>" & RS("AmeriAdjust") &"</td>"
	
	response.write "<td>" & RS("Pro20") &"</td>"
	response.write "<td>" & RS("Pro60") &"</td>"
	'response.write "<td>" & RS("ProMat") &"</td>"
	response.write "<td>" & RS("ProAdjust") &"</td>"
	
	response.write "<td>" & RS("Pertici20") &"</td>"
	response.write "<td>" & RS("Pertici60") &"</td>"
	'response.write "<td>" & RS("PerticiMat") &"</td>"
	response.write "<td>" & RS("PerticiAdjust") &"</td>"

	response.write "<td>" & RS("2Ameri20") &"</td>"
	response.write "<td>" & RS("2Ameri60") &"</td>"
	'response.write "<td>" & RS("2AmeriMat") &"</td>"
	response.write "<td>" & RS("2AmeriAdjust") &"</td>"
	
	response.write "<td>" & RS("2Pertici20") &"</td>"
	response.write "<td>" & RS("2Pertici60") &"</td>"
	'response.write "<td>" & RS("2PerticiMat") &"</td>"
	response.write "<td>" & RS("2PerticiAdjust") &"</td>"
	
	
	
	response.write "<td>" & RS("testby") & "</td>"

	response.write "<td>" & RS("Notes") & "</td>"
	response.write " </tr>"

	rs.movenext
loop
response.write "</tbody></table>"


rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing



%>
      </ul>                 
            
     
               
</body>
</html>
