<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)

part = REQUEST.QueryString("part")
supplierpart = REQUEST.QueryString("supplierpart")
des = REQUEST.QueryString("description")
kgm = REQUEST.QueryString("kgm")
if kgm = "" then
	kgm = 0 
end if
MinLevel = REQUEST.QueryString("MinLevel")
if MinLevel = "" then
	MinLevel = 250 
end if
paintcat = REQUEST.QueryString("paintcat")
CanArt = REQUEST.QueryString("CanArt")
HYDRO = REQUEST.QueryString("HYDRO")
Keymark = REQUEST.QueryString("Keymark")
Extal = REQUEST.QueryString("Extal")
inventorytype = REQUEST.QueryString("inventorytype")

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

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER"
rs.Cursortype = GetDBCursorTypeInsert
rs.Locktype = GetDBLockTypeInsert
rs.Open strSQL, DBConnection

	rs.AddNew
	rs.Fields("Part") = part
	rs.Fields("supplierpart") = supplierpart
	rs.Fields("description") = des
	rs.Fields("MinLevel") = Minlevel
	'rs.Fields("paintcat") = paintcat
	rs.Fields("CanArt") = CanArt
	rs.Fields("HYDRO") = HYDRO
	rs.Fields("Keymark") = Keymark
	rs.Fields("Extal") = Extal
	rs.Fields("inventorytype") = inventorytype
	if inventorytype = "Plastic" then
	rs.Fields("lbf") = kgm
	else
	rs.Fields("kgm") = kgm
	end if

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)

	rs.update

Call StoreID1(isSQLServer, rs.Fields("ID"))

DbCloseAll

End Function

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="masteradd.asp#_enter" target="_self">Add Master</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>


    
<ul id="Report" title="Added" selected="true">
	<li><% response.write "Q Part #: " & part %></li>
	<li><% response.write "Description: " & des %></li>
    <li><% response.write "Supplier #: " & supplierpart %></li>
    <li><% response.write "KG/M: " & kgm %></li>
	<li><% response.write "Min Stock Level: " & MinLevel %></li>
    <li><% response.write "Paint Cat.: " & paintcat %></li>
	<li><% response.write "InventoryType: " & inventorytype %></li>

</ul>



</body>
</html>

<% 

'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection=nothing
%>

