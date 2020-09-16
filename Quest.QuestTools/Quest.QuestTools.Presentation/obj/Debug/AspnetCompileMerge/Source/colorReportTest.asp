<!--#include file="dbpathTrial.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
 <!-- <script src="sorttable.js"></script>-->
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  
 
<!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>
<!-- Fixed Headers -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/extensions/FixedHeader/js/dataTables.fixedHeader.js"></script>
 
  <script type="text/javascript">
  $(document).ready( function () {
    $('#color').dataTable( {
	"pageLength": 25
     } 
);
} );
  
  </script>

<% 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [MS Access;DATABASE=f:\database\Quest.mdb].[Y_COLOR] ORDER BY PROJECT, CODE ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM [MS Access;DATABASE=f:\database\Quest.mdb].[Z_Jobs] ORDER BY JOB ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"> Color Report</h1>
		
		<% 
		
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write HomeSite%>#_Job" target="_self">Job/Colour<%response.write HomeSiteSuffix%></a>

    </div>

<ul id="screen1" title="View All Colors" selected="true">
   
    <%


response.write "<li class='group'>All Project/Colour Information </li>"
response.write "<li class='group'>" & Request.ServerVariables("REMOTE_ADDR") & "</li>"
response.write "<li>TEST: Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<p><table border='1' class='color' id ='color'><thead><tr><th>Project</th><th>Job</th><th>Parent</th><th>Ext / Int</th><th>Paint Code</th><th>Paint Type</th><th>Description</th><th>Price Category</th><th>Active</th><th>Extrusion</th><th>Sheet</th></tr></thead><tbody>"
do while not rs.eof

rs2.filter = "JOB = '" & rs.fields("JOB") & "'"
if rs2.eof then
	ParentValue = "N/A"
else
	ParentValue = rs2.fields("PARENT")
end if



response.write "<tr>"
response.write "<td>" &  rs.fields("PROJECT") & "</td>"
response.write "<td>" &  rs.fields("JOB") & "</td>"
response.write "<td>" &  ParentValue & "</td>"
response.write "<td>" &  rs.fields("SIDE") & "</td>"
response.write "<td>" &  rs.fields("CODE") & "</td>"
response.write "<td>" &  rs.fields("COMPANY") & "</td>"
response.write "<td>" &  rs.fields("DESC") & "</td>"
response.write "<td>" &  rs.fields("PRICECAT") & "</td>"
response.write "<td>" &  rs.fields("ACTIVE") & "</td>"
response.write "<td>" &  rs.fields("EXTRUSION") & "</td>"
response.write "<td>" &  rs.fields("SHEET") & "</td>"
response.write "</tr>"

rs.movenext
loop

RESPONSE.WRITE "</tbody></table></p></UL>"


rs.close
set rs=nothing
rs2.close
set rs2 = nothing

DBConnection.close
set DBConnection=nothing
%>

</body>
</html>

