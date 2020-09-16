<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
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


<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>
    
      
    
<ul id="screen1" title="Stock Level" selected="true">            


<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'SAPA' OR WAREHOUSE = 'HYDRO' OR WAREHOUSE = 'CANART' OR WAREHOUSE = 'EXTAL') order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill/Painted </li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Mill</th><th>Painted</th><th>Pending</th><th>Min level</th><th>Alerts</th></tr>"

rs2.movefirst
	do while not rs2.eof
	partqty = 0
	partqty2 = 0
	partqty3 = 0

	rs.movefirst
	do while not rs.eof
		IF RS("WAREHOUSE") = "GOREWAY" then
			IF rs2("Part") = rs("part") then
				if rs("colour") = "Mill" then
					partqty = rs("Qty") + partqty
				else
					partqty2 = rs("Qty") + partqty2
				end if
			End IF
		End if

		IF rs("WAREHOUSE") = "DURAPAINT" OR rs("WAREHOUSE") = "SAPA" OR rs("WAREHOUSE") = "HYDRO" OR RS("WAREHOUSE") = "CANART" OR RS("WAREHOUSE") = "EXTAL SEA" OR RS("WAREHOUSE") = "DEPENDABLE" then
			IF rs2("Part") = rs("part") then
				partqty3 = rs("Qty") + partqty3
			End IF
		End if

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if partqty + partqty3 < rs2("MinLevel") AND partqty2 + partqty3 < rs2("MinLevel") then
		MinLevelAlert = "Below Minimum"
	end if

	if partqty = 0 AND partqty2 = 0 then
	else
		response.write "<tr><td><a href='stockbydie.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td><td>" & partqty & "</td><td> " & partqty2 & "</td><td> " & partqty3 & "</td><td> " & rs2("MinLevel") & "</td><td> " & MinLevelAlert & "</td></tr>"
end if 

rs2.movenext
loop
response.write "</table></li>"

%>

   
            
   </ul>
</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

