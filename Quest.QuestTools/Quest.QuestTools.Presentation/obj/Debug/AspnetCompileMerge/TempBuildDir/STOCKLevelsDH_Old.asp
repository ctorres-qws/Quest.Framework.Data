<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		<!--Combined View of Stock Levels in Durapaint, Durapaint(WIP) and Horner - As per Request of Shaun Levy  -->
		<!--Michael Bernholtz January 2015--> 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels - Mill</title>
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
    
      
    
<ul id="screen1" title="Durapaint and Horner" selected="true">            
<li class="group"><a href="STOCKLevelsDurapaint.asp" target="_self" >GO TO Stock Levels Durapaint (Inventory and WIP) Only</a></li>
<li class="group"><a href="STOCKLevelsHorner.asp" target="_self" >GO TO Stock Levels Horner Only</a></li>
<li class="group"><a href="STOCKLevelsTilton.asp" target="_self" >GO TO Stock Levels Tilton Only</a></li>
<li class="group"><a href="STOCKLevelsTorbram.asp" target="_self" >GO TO Stock Levels Torbram Only</a></li>

<%
Set rs = Server.CreateObject("adodb.recordset")

If b_SQL_Server Then

strSQL = strSQL & "SELECT T.*, T.PartQty+T.AllocateQty+T.PaintQty as TotalQty FROM "
strSQL = strSQL & "( "
strSQL = strSQL & "SELECT y_M.Part, y_M.Description, "
strSQL = strSQL & "SUM(CASE WHEN y_I.Colour = 'MILL' AND RTRIM(LTRIM(ISNULL(Allocation,''))) = '' THEN y_I.Qty ELSE 0 END) as PartQty, "
strSQL = strSQL & "SUM(CASE WHEN y_I.Colour = 'MILL' AND RTRIM(LTRIM(ISNULL(Allocation,''))) <> '' THEN y_I.Qty ELSE 0 END) as AllocateQty, "
strSQL = strSQL & "SUM(CASE WHEN ISNULL(y_I.Colour,'') <> 'MILL' THEN y_I.Qty ELSE 0 END) as PaintQty "
strSQL = strSQL & "FROM Y_MASTER y_M "
strSQL = strSQL & "INNER JOIN y_INV y_I on y_I.Part = y_M.part "
strSQL = strSQL & "WHERE WAREHOUSE = 'DURAPAINT'  OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TILTON'  OR WAREHOUSE = 'TILTON(WIP)' OR WAREHOUSE = 'TORBRAM' "
strSQL = strSQL & "GROUP BY y_M.Part, y_M.Description "
strSQL = strSQL & ") T "
strSQL = strSQL & "ORDER BY PART ASC "

'Get a Record Set
    Set RS2 = DBConnection.Execute(strSQL)

response.write "<li>Stock by Die - Durapaint, Durapaint(WIP), NASHUA, Torbram, Tilton and Horner </li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Unallocated Mill</th><th>Allocated Mill</th><th>Painted Stock</th><th>Total Qty</th></tr>"

rs2.movefirst
do while not rs2.eof
	partqty = rs2("PartQty")
	allocatedqty = rs2("AllocateQty")
	Totalqty = rs2("TotalQty")
	paintedqty = rs2("PaintQty")

	if partqty = 0 and allocatedqty = 0 and paintedqty = 0 then
	else
		response.write "<tr><td><a href='stockbydieDH.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td><td>" & rs2("Description") & "</td><td>" & partqty & "</td><td>" & allocatedqty & "</td><td>" & paintedqty & "</td><td>" & Totalqty & "</td></tr>"
	end if 

rs2.movenext
loop
response.write "</table></li>"

Else

strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'DURAPAINT'  OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TILTON'  OR WAREHOUSE = 'TILTON(WIP)' OR WAREHOUSE = 'TORBRAM' order by PART ASC"

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li>Stock by Die - Durapaint, Durapaint(WIP), NASHUA, Torbram, Tilton and Horner </li>"
response.write "<li><table border='1' class='sortable'><tr><th>Stock</th><th>Description</th><th>Unallocated Mill</th><th>Allocated Mill</th><th>Painted Stock</th><th>Total Qty</th></tr>"

rs2.movefirst
do while not rs2.eof
	partqty = 0
	allocatedqty = 0
	Totalqty = 0
	paintedqty = 0

	rs.movefirst
	do while not rs.eof
		IF rs2("Part") = rs("part") then
		
		Totalqty = rs("Qty") + Totalqty
			if rs("colour") = "Mill" then
				if rs("Allocation") = "" OR isNUll(rs("Allocation")) then
					partqty = rs("Qty") + partqty
				else
					allocatedqty = rs("Qty") + allocatedqty
				End if
			else
				paintedqty = rs("Qty") + paintedqty
			End if
		End if
	rs.movenext
	loop

	if partqty = 0 and allocatedqty = 0 and paintedqty = 0 then
	else
		response.write "<tr><td><a href='stockbydieDH.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td><td>" & rs2("Description") & "</td><td>" & partqty & "</td><td>" & allocatedqty & "</td><td>" & paintedqty & "</td><td>" & Totalqty & "</td></tr>"
	end if 

rs2.movenext
loop
response.write "</table></li>"


End If
%>
<li><a href="stockpendingtable.asp?part=<%response.write part%>" target="_self" >View Stock Pending</a></li>
   
            
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

