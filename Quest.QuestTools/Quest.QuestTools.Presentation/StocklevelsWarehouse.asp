<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 <!--Created May 23rd, by Michael Bernholtz - Overarching tool-->
		 <!--All  Warehouse version of Stock levels -->
		 <!-- Unsure if this will be a production tool, currently not in Use-->
		

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels - Durapaint</title>
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
                <a class="button leftButton" type="cancel" href="StockLevelsWarehouse1.asp" target="_self">Search</a>
    </div>
    
 <%     
Warehouse = Request.Querystring("Warehouse")    
response.write "<ul id='screen1' title='Stock Level - " & WareHouse & "' selected='true'>"            



Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = '" & WareHouse & "' order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection




'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - " & WareHouse & " Only</li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Qty</th></tr>"

if rs.eof then
Response.write "<tr><td colspan ='2'> No Stock </td>"
else



rs2.movefirst
do while not rs2.eof
	partqty = 0

	rs.movefirst
	do while not rs.eof
		IF rs2("Part") = rs("part") then
			partqty = rs("Qty") + partqty
		End if
	rs.movenext
	loop

	if partqty = 0 then
	else
		response.write "<tr><td><a href='stockbydie.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td><td>" & partqty & "</td></tr>"
	end if 

rs2.movenext
loop

end if

response.write "</table></li>"



%>


<% 
if Warehouse = "SAPA" or Warehouse = "HYDRO" or Warehouse = "Durapaint" or Warehouse = "EXTAL SEA" then
%>
<li><a href="stockpendingtable.asp?part=<%response.write part%>" target="_self" >View Stock Pending</a></li>
<%
end if 
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

