<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Virtual Directory - F:\projects\15 - Inventory Tracking\die's pick date base-->
<!-- New Directory - \\172.18.13.31\_Websites\Prod\QWS_Tools\partpic -->
<!-- February 2019 - Added USA view -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
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
    $('#Extrusion').DataTable();
} );
  $(document).ready( function () {
    $('#Gasket').DataTable();
} );
  $(document).ready( function () {
    $('#Hardware').DataTable();
} );
  $(document).ready( function () {
    $('#Plastic').DataTable();
} );
  $(document).ready( function () {
    $('#Sheet').DataTable();
} );
    $(document).ready( function () {
    $('#Other').DataTable();
} );
  </script>
    <script type="text/javascript">

  
  </script>
  
  <style>
body{
zoom: 90%
};
 </style>
  
 <% 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER ORDER BY Inventorytype, Part"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



sortby = REQUEST.QueryString("sortby")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Stock by Picture</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>

<ul id="screen1" title="View Stock - Part List" selected="true">
    
    <li class="group">Extrusion List</li>
    <%
	
	rs.filter = "inventorytype = 'Extrusion'"

RESPONSE.WRITE "<li><table border='1' class='Extrusion' id = 'Extrusion'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>HYDRO</th><th>Picture</th><th>KGM</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof


Response.write "<tr>"
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td> " & rs("HYDRO") & "</td>"	
Response.write "<td><img height ='200' width = '200' src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"
%>
    <li class="group">Gasket List</li>
    <%
	
	rs.filter = "inventorytype = 'Gasket'"

RESPONSE.WRITE "<li><table border='1' class='Gasket' id = 'Gasket'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>Picture</th><th>KGM</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<tr>"
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"

%>

    <li class="group">Hardware List</li>
    <%
	
	rs.filter = "inventorytype = 'Hardware'"

RESPONSE.WRITE "<li><table border='1' class='Hardware' id = 'Hardware'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>Picture</th><th>KGM</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<tr>"
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"

%>

    <li class="group">Plastic List</li>
    <%
	
	rs.filter = "inventorytype = 'Plastic'"
	
RESPONSE.WRITE "<li><table border='1' class='Plastic' id = 'Plastic'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>Picture</th><th>LBF</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<tr>"
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("LBF") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"

%>

    <li class="group">Sheet List</li>
    <%
	
	rs.filter = "inventorytype = 'Sheet'"

RESPONSE.WRITE "<li><table border='1' class='Sheet' id = 'Sheet'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>Picture</th><th>KGM</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof
	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<tr>"
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"

%>

    <li class="group">All Other items List</li>
    <%
	
	rs.filter = "inventorytype = NULL"
	
RESPONSE.WRITE "<li><table border='1' class='Other' id = 'Other'><thead>"
RESPONSE.WRITE "<tr><th>View Stock</th><th>Part</th><th>Picture</th><th>KGM</th><th>Description</th></tr></thead><tbody>"

do while not rs.eof

	part = rs("part")
	partfilename = part
		if instr(1, partfilename, chr(47))>0 then
		partfilename = replace (partfilename, chr(47), "-")
	end if
Response.write "<tr>"
Response.write "<td><a href='stockbydie.asp?part=" & part & "&ticket=pic' target='_self'>View Stock</a></td>"
Response.write "<td> " & rs("PART") & "</td>"
Response.write "<td><img src='/partpic/" & partfilename & ".png'/></td>"
Response.write "<td> " & rs("KGM") & "</td>"
Response.write "<td> " & rs("Description") & "</td>"

Response.write "</tr>"
	
rs.movenext
loop
Response.write "</tbody></table></li>"




RESPONSE.WRITE "</UL>"


%>

</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>
