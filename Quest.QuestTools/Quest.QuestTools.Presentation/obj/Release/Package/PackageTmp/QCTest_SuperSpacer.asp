                      
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->

<!-- Testing Results stored in the system - Designed for Victor Babuskins - November 2014, Michael Bernholtz-->
<!-- Main Page -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Super Spacer Test</title>
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
strSQL = "SELECT * FROM TEST_SuperSpacer ORDER BY DATE DESC"
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
      
       
        <ul id="Profiles" title="SuperSpacer" selected="true">
        <li class='group'>Super Spacer Test </li>
         <a class="whiteButton" href="QCTest_SuperSpacerEnter.asp" target='_Self'>Add New Result</a>
<% 

response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable' ><thead><tr><th>Date</th><th>Time</th><th>Condition</th><th>Width</th><th>Adhesion Passed?</th><th>Name</th><th>Notes</th><th>File</th></tr></thead><tbody>"
if rs.eof then
Response.write "<tr><td colspan ='14'>No current orders</td></tr>"
end if	
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("Date") & "</td>"
	response.write "<td>" & RS("Time") &"</td>"
	response.write "<td>" & RS("Condition") &"</td>"
	response.write "<td>" & RS("Width") &"</td>"
	response.write "<td>" & RS("Adhesion") & "</td>"
	response.write "<td>" & RS("Initials") & "</td>"
	response.write "<td>" & RS("Notes") & "</td>"
	response.write "<td>" & RS("File") & "</td>"
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
