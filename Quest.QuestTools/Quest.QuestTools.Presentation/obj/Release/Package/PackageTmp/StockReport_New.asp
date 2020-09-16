<!--#include file="dbpath.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!-- Limited to Extrusion Only, Seperate report for Plastic - Michael Bernholtz, at request of Shaun Levy December 2014-->
 <script src="sorttable.js"></script>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'DURAPAINT(WIP)' order by PART ASC"




'Create a Query
    strSQL = "SELECT YI.*,IIF(YM.InventoryType = 'Extrusion',YM.kgm,-1) as kgm, YM.InventoryType FROM Y_MASTER YM "
    strSQL = strSQL & " INNER JOIN Y_INV YI ON YI.Part = YM.Part "
    strSQL = strSQL & " WHERE YI.WAREHOUSE = 'GOREWAY' OR YI.WAREHOUSE = 'NASHUA' OR YI.WAREHOUSE = 'DURAPAINT' OR YI.WAREHOUSE = 'HORNER' OR YI.WAREHOUSE = 'DURAPAINT(WIP)' "
    strSQL = strSQL & " order BY YI.Part asc"

rs.Cursortype = 2           
rs.Locktype = 3             
rs.Open strSQL, DBConnection

Response.Write(strSQL)

'Get a Record Set
    'Set RS2 = DBConnection.Execute(SQL2)
'Aluminum price
alumprice = 3.95

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<P> Reporting Value of Inventory in Goreway, Nashua, Horner, Durapaint, Durapaint(WIP) </P>
<P> Que 157 Price changed to 0.35 </P>
<table border='1' class='sortable'>
  <tr><th>Part</th>
    <th>Qty</th>
    <th>Length (mm)</th>
    <th>Colour</th>
	 <th>Bundle</th>
   <th>KGM</th>
    <th>Value</th>
    </tr>
<%
Server.ScriptTimeout=400
%>
<%

rs.movefirst
do while not rs.eof
invpart = rs("part")
partqty = rs("Qty")
lmm = CDBL(rs("lmm"))
linch = CDBL(rs("linch"))
colour = rs("colour")
project = rs("project")
bundle = rs("bundle")
errormsg = 1
NotEx = 1

If rs("kgm") <> "-1" Then
	kgm = rs("kgm")
	errormsg = 0
	NotEx = 0
End If
'		NotEx = 0       
'rs2.movefirst
'do while not rs2.eof
'if rs2("inventoryType") = "Extrusion" then
'	else
'		if NotEX = 0 then
'		else
'		NotEx = 1
'		end if
'	end if
'rs2.movenext
'loop
if errormsg = 1 and NotEx =0 then
response.write invpart & " not in inventory master - " & rs("id") & " <BR>"
end if
if NotEx =  0 then
'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

	if kgm > 5 then 
	pricebar = partqty * kgm
	value2 = value2 + pricebar
	'this code is 5 times overvalued
	else
		if invpart = "Que-157" then
			tempvalue =  0.35 * (partqty * (kgm * ( lmm / 1000 )))
		else
			tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
		end if
	value = value + tempvalue
	end if

	if colour = "White" Then
	paintlft = (linch/12) * partqty
	totalpaintlft = totalpaintlft + paintlft
	else
	
	If colour = "Mill" then
	else
	paintlft2 = (linch/12) * partqty
	totalpaintlft2 = totalpaintlft2 + paintlft2
	end if
	
	end if

%>

  <tr><td><% response.write invpart %></td>
    <td><% response.write partqty %></td>
    <td><% response.write lmm %></td>
    <td><% response.write colour %></td>
	<td><% response.write bundle %></td>
    <td>$<% response.write round(KGM, 3) %></td>
    <td>$<% if kgm > 5 then
	response.write round(pricebar,2)
	else
	response.write round(tempvalue,2)
	end if %></td>
    </tr>

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

	end if ' End NotEx code

rs.movenext
loop

%> </Table> <%

paintvalue1 = (totalpaintlft *0.14)
paintvalue2 = (totalpaintlft2 *0.38)
response.write "<BR><HR>"
response.write "Total kgm Material $" & round(value,2) & "<BR>"
response.write "Total perbar Material $" & round(value2,2)  & "<BR>"
response.write "<HR>"
response.write "SubTotal = $" & round(value,2) + round(value2,2) & "<BR><BR>"
response.write "Total Paint White $" & paintvalue1 & "<BR>"
response.write "Total Paint Project $" & paintvalue2 & "<BR>"
response.write "<HR>"
response.write "Grand Total = $" & round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2)

%>

       
            
    
</body>
</html>

<% 

rs.close
set rs=nothing
'rs2.close
'set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

