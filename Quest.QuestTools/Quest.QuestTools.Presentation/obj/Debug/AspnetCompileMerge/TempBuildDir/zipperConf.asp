                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- Form Created December 31, 2014, Michael Bernholtz at request of Slava Kotek, Lev Bedoev, Jody Cash-->
<!-- Entry Form to add items to Zipper-->
<!-- Zipper will be changed to automatic only-->
<!-- Entry from Quest Dashboard ZipperEnter.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Zipper / Rolling Extrusion</title>

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
strSQL = "SELECT * FROM Roll_TABLE ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
if rs.eof then 
ID =1
else
rs.movelast
ID = Rs("ID") +1
end if


currentDate = Date()

job = request.QueryString("Job")
floor = request.QueryString("Floor")
MullionQty = request.Querystring("MullionQty")
if MullionQty = "" then
MullionQty = 0
end if
SashQty = request.Querystring("SashQty")
if SashQty = "" then
SashQty = 0
end if
SillQty = request.Querystring("SillQty")
if SillQty = "" then
SillQty = 0
end if
JambQty = request.Querystring("JambQty")
if JambQty = "" then
JambQty = 0
end if

MullionLft = request.Querystring("MullionLft")
if MullionLft = "" then
MullionLft = 0
end if
SashLft = request.Querystring("SashLft")
if SashLft = "" then
SashLft = 0
end if
SillLft = request.Querystring("SillLft")
if SillLft = "" then
SillLft = 0
end if
JambLft = request.Querystring("JambLft")
if JambLft = "" then
JambLft = 0
end if

sheartest = request.QueryString("sheartest")
if UCASE(sheartest) = "YES" or  UCASE(sheartest) = "YES" or sheartest = "" then
sheartest = "Yes"
else
sheartest = "No"
end if

EntryCount = 0

if MullionQty < 1 then
else
rs.AddNew
	rs.Fields("ID") = ID
	ID = ID + 1
	rs.Fields("Job") = Job
	rs.Fields("Floor") = Floor
	rs.Fields("qty") = Round(MullionQty,0)
	rs.Fields("enterDate") = currentDate
	rs.Fields("length") = MullionLft
	rs.Fields("sheartest") = sheartest
	rs.Fields("Profile") = "MULLION"
	EntryCount = EntryCount + 1
	rs.update
end if

if SashQty < 1 then
else
rs.AddNew
	rs.Fields("ID") = ID
	ID = ID+1
	rs.Fields("Job") = Job
	rs.Fields("Floor") = Floor
	rs.Fields("qty") = Round(SashQty,0)
	rs.Fields("enterDate") = currentDate
	rs.Fields("length") = SashLft
	rs.Fields("sheartest") = sheartest
	rs.Fields("Profile") = "SASH"
	EntryCount = EntryCount + 1
	rs.update
end if

if SillQty < 1 then
else
rs.AddNew
	rs.Fields("ID") = ID
	ID = ID+1
	rs.Fields("Job") = Job
	rs.Fields("Floor") = Floor
	rs.Fields("qty") = Round(SillQty,0)
	rs.Fields("enterDate") = currentDate
	rs.Fields("length") = SillLft
	rs.Fields("sheartest") = sheartest
	rs.Fields("Profile") = "SILL"
	EntryCount = EntryCount + 1
	rs.update
end if

if JambQty < 1 then
else
rs.AddNew
	rs.Fields("ID") = ID
	ID = ID+1
	rs.Fields("Job") = Job
	rs.Fields("Floor") = Floor
	rs.Fields("qty") = Round(JambQty,0)
	rs.Fields("enterDate") = currentDate
	rs.Fields("length") = JambLft
	rs.Fields("sheartest") = sheartest
	rs.Fields("Profile") = "JAMB/HEAD"
	EntryCount = EntryCount + 1
	rs.update
end if

rs.close
set rs = nothing
DBConnection.close
set DBConnection = nothing
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="Index.html#_Zipper" target="_self">Zipper</a>
    </div>


    
<ul id="Report" title="Zipper Request Sent to Rollers" selected="true">
<% if EntryCount = 0 then
%>
	<li>No Entries - All values are 0 or Invalid</li>
<%	else
%>
	

	<li><% response.write "Job " & job %></li>
	<li><% response.write "Floor " & floor %></li>
	<li><% response.write "Mullion: " & MullionQty & " at " & MullionLft & "feet "%></li>
	<li><% response.write "Sill: " & SillQty & " at " & SillLft & "feet "%></li>
	<li><% response.write "Sash: " & SashQty & " at " & SashLft & "feet "%></li>
	<li><% response.write "Jamb / Header: " & JambQty & " at " & JambLft & "feet "%></li>
    <li><% response.write "Counted to Shear Test (Default Yes) " & sheartest %></li>
<% End If%>
   
<li><a class = 'whiteButton' href='ZipperEnter.asp#_enter' target='_self'>Add Job/Floor </a></li>
<li><a class = 'whiteButton' href='Index.html#_Zipper' target='_self'>Zipper</a></li>

</ul>

</body>
</html>



