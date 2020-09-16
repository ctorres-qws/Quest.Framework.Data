<!--#include file="dbpath.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!-- Limited to Sheet Only, Seperate report for Plastic/Extrusion - Michael Bernholtz, at request of Shaun Levy November 2015-->
 <script src="sorttable.js"></script>
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT I.*, M.* FROM Y_INV AS I Inner Join Y_MASTER As M on I.Part = M.Part WHERE (I.WAREHOUSE = 'GOREWAY' OR I.WAREHOUSE = 'NASHUA') AND M.InventoryType = 'Sheet' order by I.PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Aluminum price
'alumprice = 3.95

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->

<table border='1' class='sortable'>
  <tr><th>Part</th>
    <th>Description</th>
	<th>Qty</th>
    <th>Thickness</th>
    <th>Colour</th>
	<th>Size</th>   
    <th>Area</th>
	<th>Bundle</th>
	<th>LBF</th>
    <th>Value</th>
    </tr>

<%

rs.movefirst
do while not rs.eof
invpart = rs("Part")
partqty = rs("Qty")
thickness = rs("Thickness")
colour = rs("colour")
project = rs("project")
bundle = rs("bundle")
width = int(rs("width"))
Height = int(rs("height"))
size = rs("width") & " by " & rs("height")
area = Int(width * height)

if len(rs("Description")) >0 then
lbf = rs("lbf")
Description = rs("Description")
else
NotSh =0
end if

if errormsg = 1 and NotSh =0 then
response.write invpart & " not in inventory master - " & rs("id") & " <BR>"
end if
if NotSh =  0 then
'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

	if lbf > 5 then 
	pricebar = partqty * lbf
	value2 = value2 + pricebar
	else
			tempvalue = (partqty * (lbf * lft ))
	value = value + tempvalue
	end if


%>

	<tr>
	<td><% response.write invpart %></td>
	<td><% response.write description %></td>
    <td><% response.write partqty %></td>  
	<td><% response.write thickness %></td>
    <td><% response.write colour %></td>
	<td><% response.write size %></td>
	<td><% response.write area %></td>
	<td><% response.write bundle %></td>
    <td>$<% response.write lbf%></td> 
    <td>$<% if lbf > 5 then
	response.write round(pricebar,2)
	else
	response.write round(tempvalue,2)
	end if %></td>
    </tr>

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
	Response.write "4"


Response.write "3"
END IF

	end if ' End NotSh code

rs.movenext
loop

%> </Table> <%
' Sheets like Extrusion have colour, Unlike Plastic
paintvalue1 = (totalpaintlft *0.14)
paintvalue2 = (totalpaintlft2 *0.38)

response.write "<BR><HR>"
response.write "Total lbf Material $" & round(value,2) & "<BR>"
response.write "Total perbar Material $" & round(value2,2)  & "<BR>"
response.write "<HR>"
response.write "SubTotal = $" & round(value,2) + round(value2,2) & "<BR><BR>"
response.write "Total Paint White $" & paintv & "<BR>"
response.write "Total Paint Project $" & paintvalue2 & "<BR>"
response.write "<HR>"
response.write "Grand Total = $" & round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2)

%>

       
            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

