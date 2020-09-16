<!--#include file="dbpath.asp"-->
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INVQ12013 WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'DURAPAINT') order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

alumprice = 3.61

'response.buffer = false
%> <Table> 
  <tr><td>Part</td>
    <td>Qty</td>
    <td>Length (mm)</td>
    <td>Colour</td>
    <td></td>
    <td>Value</td>
    </tr>

<%

rs.movefirst
do while not rs.eof
invpart = rs("part")
partqty = rs("Qty")
lmm = rs("lmm")
linch = rs("linch")
colour = rs("colour")
project = rs("project")
errormsg = 1

rs2.movefirst
do while not rs2.eof
if rs2("part") = invpart then 
kgm = rs2("kgm")
errormsg = 0
'Response.write rs2("kgm") & "<BR>"
else
end if
rs2.movenext
loop

if errormsg = 1 then
response.write invpart & " not in inventory master <BR>"
end if

'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

	if kgm > 5 then 
	pricebar = partqty * kgm
	value2 = value2 + pricebar
	'this code is 5 times overvalued
	else
	tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
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
    <td></td>
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

	

rs.movenext
loop

%> </Table> <%

paintvalue1 = (totalpaintlft *0.18)
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
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

