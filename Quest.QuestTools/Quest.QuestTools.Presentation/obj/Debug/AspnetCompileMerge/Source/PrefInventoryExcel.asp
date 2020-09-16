<!--#include file="dbpath.asp"-->
              <!-- PREF Inventory Dump Program runs the whole Database looking for Durapaint / Horner / Goreway to create a full PREF inventory list -->
			  <!-- Excel File can be created using PREFINVENTORYEXCEL.asp -->
			  <!-- July 2015 For Peter Tiede at BEST, Programmed by Michael Bernholtz-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Drill Down </title>
 

  


<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body onload="startTime()" >
<%
ticket=Request.Querystring("ticket")

%>

    <div class="toolbar">
        <a id="backButton" class="button" href="#"></a>

        
    </div>
    
      
  
<%
Set rs = Server.CreateObject("adodb.recordset")
	strSQL = " SELECT INV.WAREHOUSE, INV.PART, INV.COLOUR, INV.QTY, INV.LFT, INV.PREF, MASTER.PART, MASTER.INVENTORYTYPE, COLOR.PROJECT, COLOR.CODE FROM Y_INV AS INV, Y_MASTER AS MASTER, Y_COLOR AS COLOR WHERE INV.PART = MASTER.PART AND INV.COLOUR = COLOR.PROJECT AND (INV.Warehouse = 'GOREWAY' OR INV.Warehouse = 'HORNER' OR INV.Warehouse = 'DURAPAINT' OR INV.Warehouse = 'DURAPAINT(WIP)' OR INV.Warehouse = 'MILVAN' OR INV.Warehouse = 'TORBRAM'  OR INV.WAREHOUSE = 'NASHUA' OR INV.Warehouse = 'TILTON' OR WAREHOUSE = 'NPREP') AND MASTER.InventoryType ='Extrusion' order by INV.PART ASC, COLOR.CODE ASC, INV.Lft ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

TodayDay = Month(Now) & "_" & Day(Now)
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=QuestInventory" & TodayDay & ".xls"
%>


<ul id="screen1" selected="true">            
<li><table border='1' class='sortable' ><tr><th>Part</th><th>Colour</th><th>PREF</th><th>Length</th><th>Count</th></tr>
<%


Part = ""
Colour = ""
Length = ""
COUNTQTY = 0

Do while Not rs.eof

PrePart = Part
PreColour = Colour
preLength = Length
PrePref = Pref
preQty = QTY
Part = rs("PART")
Colour = rs("CODE")
Length = rs("Lft")
QTY = rs("QTY")
Pref = rs("PREF")
CountQTY = COUNTQTY
if PART = PREPart then
	if Colour = PreColour then
		if CINT(Length) = CINT(PreLength) then	
		CountQTY = CountQTY + QTY
		else 
			Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
			COUNTQTY =  QTY
		end if 
	else 
		Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
		COUNTQTY =  QTY
	end if 
else 	
	if PREPART = "" then
	else
	Response.write "<tr><td>" & PREPART & "</td><td>" & PREColour & "</td><td>" & PrePref & "</td><td>" & PRELength & "</td><td>" & COUNTQTY & "</td>"
	end if
	COUNTQTY =  QTY
	
end if 
		

	rs.movenext
	loop
Response.write "<tr><td>" & PART & "</td><td>" & Colour & "</td><td>" & PrePref & "</td><td>" & Length & "</td><td>" & COUNTQTY & "</td>"


rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
%>

   
            
   </ul>
</body>
</html>



