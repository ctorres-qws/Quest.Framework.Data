                       
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
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE [HIDE] IS NULL AND [DEPARTMENT] = 'Service'  ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_Master ORDER BY ID ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM Y_COLOR ORDER BY ID ASC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "SELECT * FROM Y_INV ORDER BY ID ASC"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection


DO while not rs4.eof
PREFNAME =""
rs2.filter = "PART = '" & rs4("Part") & "'"
if rs2.eof then
PREFNAME = x
else
PREFNAME = rs2("PREFREF")
end if
rs3.filter =" PROJECT = '" & rs4("Colour")& "'"
if rs3.eof then
PREFNAME = PREFNAME & " x"
else
PREFNAME = PREFNAME & " " & rs3("CODE")
end if

rs4("PREF") =PREFNAME
rs4.update
response.write PREFNAME
response.write "<br>"
rs4.movenext
loop
rs2.close
rs3.close
rs4.close
set rs2 = nothing
set rs3 = nothing
set rs4 = nothing
'afilter = request.QueryString("aisle")


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
response.write "<li class='group'>SERVICE GLASS REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>Notes</th><th>Status</th><th>TimeLine</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td><td>" & RS("NOTES") & "</td>"
	%>
		<!--#include file="GlassStatus.inc"-->
	<%
	Response.write "<td>" & Status & "</td> " ' NStatus
	response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=service' target ='#_blank' >Time Line</a> </td>"
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li></ul>"



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
