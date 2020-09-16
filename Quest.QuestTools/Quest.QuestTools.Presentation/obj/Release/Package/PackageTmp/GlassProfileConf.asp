<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Glass Profile Table Updates GlassTypes Table for Glass Entry - Job Specific on 8080-->


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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>


<% 

NAME = Request.Form("NAME")
DESCRIPTION = Request.Form("DESCRIPTION")
JOB = Request.Form("JOB")
MATERIAL = Request.Form("MATERIAL")
ExtGlass = Request.Form("ExtGlass")
IntGlass = Request.Form("IntGlass")
ExtGlassDoor = Request.Form("ExtGlassDoor")
IntGlassDoor = Request.Form("IntGlassDoor")
FixWindowThick = Request.Form("FixWindowThick")
SwingDoorThick = Request.Form("SwingDoorThick")
CasAwnThick = Request.Form("CasAwnThick")
SUSpacer = Request.Form("SUSpacer")
OVSpacer = Request.Form("OVSpacer")
SpacerColour = Request.Form("SpacerColour")
Silltype = Request.Form("Silltype")
Gas = Request.Form("Gas")


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "INSERT INTO GlassTypes ("
strSQL = strSQL & "[NAME], [DESCRIPTION], [JOB], [MATERIAL],  [ExtGlass],[IntGlass], [ExtGlassDoor], [IntGlassDoor], [FixWindowThick], [SwingDoorThick], " 
strSQL = strSQL & "[CasAwnThick], [SUSpacer],[OVSpacer],[SpacerColour],[Silltype],[Gas]" 
strSQL = strSQL & " Values "

strSQL = strSQL & " ('" & Name & "', '" & Description &  "', '" & JOB & "', '" & Material &  "', '" & ExtGlass &  "', '" & IntGlass & "', '" & ExtGlassDoor &  "', '" & IntGlassDoor & "', '" & FixWindowThick & "', '" & SwingDoorThick & "', "
strSQL = strSQL & " '" & CasAwnThick & "', '" & SUSpacer & "', '" & OVSpacer & "', '" &SpacerColour & "', '" & Silltype & "', '" & Gas & "') "

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


	
%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassProfileEnter.asp" target="_self">Panel Entry</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<!-- Continue here-->
    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Name " & Name %></li>
	<li><% response.write "Description " & Description %></li>
    <li><% response.write "Job " & Job %></li>
	<li><% response.write "Side " & Side %></li>
    <li><% response.write "Material " & Material %></li>
    <li><% response.write "Colour " & Colour %></li>
    <li><% response.write "Offset X " & OffsetX %></li>
	<li><% response.write "Offset Y " & OffsetY %></li>
	<li><% response.write "Notes " & Notes %></li>
	<li><% response.write "Side " & Side %></li>
    <li><% response.write "Material " & Material %></li>
    <li><% response.write "Colour " & Colour %></li>
    <li><% response.write "Offset X " & OffsetX %></li>
	<li><% response.write "Offset Y " & OffsetY %></li>
	<li><% response.write "Notes " & Notes %></li>

	</ul>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
%>

</body>
</html>



