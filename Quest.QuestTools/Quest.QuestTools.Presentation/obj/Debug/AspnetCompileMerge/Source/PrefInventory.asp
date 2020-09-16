<!--#include file="dbpath.asp"-->
              <!-- PREF Inventory Dump Program runs the whole Database looking for Durapaint / Horner / Goreway to create a full PREF inventory list -->
			  <!-- Excel File can be created using PREFINVENTORYEXCEL.asp -->
			  <!-- July 2015 For Peter Tiede at BEST, Programmed by Michael Bernholtz-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Drill Down </title>
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
<%
ticket=Request.Querystring("ticket")

%>

    <div class="toolbar">
        <h1 id="pageTitle">Full Inventory</h1>
        <a id="backButton" class="button" href="#"></a>
		  <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
        
    </div>

<%

Set rs = Server.CreateObject("adodb.recordset")
	strSQL = " SELECT INV.WAREHOUSE, INV.PART, INV.COLOUR, INV.QTY, INV.LFT, INV.PREF, MASTER.PART, MASTER.INVENTORYTYPE, COLOR.PROJECT, COLOR.CODE FROM Y_INV AS INV, Y_MASTER AS MASTER, Y_COLOR AS COLOR WHERE INV.PART = MASTER.PART AND INV.COLOUR = COLOR.PROJECT AND INV.Warehouse IN ('GOREWAY','HORNER','MILVAN','NASHUA','DURAPAINT','DURAPAINT(WIP)','NPREP') AND MASTER.InventoryType ='Extrusion' ORDER BY INV.PART ASC, COLOR.CODE ASC, INV.Lft ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL = "SELECT * FROM Y_INV where Warehouse = 'GOREWAY' OR Warehouse = 'HORNER' OR Warehouse = 'DURAPAINT' OR Warehouse ='NASHUA' OR Warehouse = 'DURAPAINT(WIP)' OR Warehouse = 'TILTON' OR Warehouse = 'TORBRAM)' ORDER BY ID ASC"
strSQL = "SELECT yI.*, yM.PrefRef, yC.Code FROM ((Y_INV yI LEFT JOIN y_Master yM ON yM.Part = yI.Part) LEFT JOIN y_Color yC ON yC.Project = yI.Colour) WHERE yI.Warehouse IN('GOREWAY','HORNER','DURAPAINT','NASHUA','DURAPAINT(WIP)','TILTON','MILVAN','TORBRAM)','NPREP') AND Pref IS NULL ORDER BY yI.ID ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL, DBConnection

Do while not rs2.eof
	part = rs2("part")
	colour = rs2("colour")

	If rs2("PrefRef") & "" <> "" Then
		PREFName = RS2("PREFREF")
	Else
		PREFName = "x"
	End If

	If rs2("Code") & "" <> "" Then
		PREFColour = rs2("CODE")
	Else
		PREFColour = "x"
	End If

	PREFValue = PREFName & " " & PREFColour

	DBConnection.Execute("Update y_Inv SET Pref='" & PREFValue & "' WHERE ID=" & rs2("ID"))

	rs2.MoveNext
loop

rs2.close
set rs2 = nothing

%>

<ul id="screen1" title="Full Inventory " selected="true">     
<li class="group"><a href="PrefInventoryexcel.asp" target="_self" >Download Excel File</a></li>       
<li><table border='1' class='sortable' ><tr><th>Part</th><th>Colour</th><th>PREF</th><th>Length</th><th>Count</th></tr>
<%

	Part = ""
	Colour = ""
	Length = ""
	COUNTQTY = 0

	Do while Not rs.eof
		PrePart = Part
		PreColour = Colour
		PreLength = Length
		PrePref = Pref
		preQty = QTY
		Part = rs("PART")
		Colour = rs("CODE")
		Length = rs("Lft")
		QTY = rs("QTY")
		Pref = rs("PREF")
		CountQTY = COUNTQTY

		If PART = PREPart Then
			If Colour = PreColour Then
				If CINT(Length) = CINT(PreLength) Then
					CountQTY = CountQTY + QTY
				Else 
					Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
					COUNTQTY =  QTY
				End If
			Else
				Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
				COUNTQTY =  QTY
			End If
		Else
			If PREPART = "" Then
			Else
				Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
			End If
			COUNTQTY =  QTY
		End If

		rs.movenext
	Loop

	Response.write "<tr><td>" & PART & "</td><td>" & Colour & "</td><td>" & PrePref & "</td><td>" & Length & "</td><td>" & COUNTQTY & "</td>"

	rs.close
	set rs=nothing
	DBConnection.close
	set DBConnection=nothing
%>

   </ul>
</body>
</html>