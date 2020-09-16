<!--#include file="dbpath.asp"-->
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INVQ12013 WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'DURAPAINT') order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

alumprice = 3.68
newvalue = 0

'response.buffer = false

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
	end if
	
	If not IsNull(Project) then
	paintlft2 = (linch/12) * partqty
	totalpaintlft2 = totalpaintlft2 + paintlft2
	end if

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

if kgm > 5 then 
	pricebar = partqty * kgm
	value = pricebar
	'response.write rs("ID") & " " & rs("Part") & " " & partqty & " XXXX " & kgm & " XXX " & value & "<BR>" 
	'this code is 5 times overvalued, hu
	else
	value =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
	'response.write rs("ID") & " " & rs("Part") & " " & partqty & " XXXX " & kgm & " XXX " & (lmm/1000) & " " & value 
	
end if
	
if kgm = 0 then
response.write rs("Part") & "has zero value in Master"
end if
	
contvalue = contvalue + value
response.write " " & contvalue & "<BR>" 

rs.movenext
loop

paintvalue1 = (totalpaintlft *0.18)
paintvalue2 = (totalpaintlft2 *0.28)
response.write "Total kgm Material $" & value & "<BR>"
response.write "Total perbar Material $" & value2  & "<BR>"
response.write "<HR>"
response.write "SubTotal = $" & value + value2 & "<BR><BR>"
response.write "Total Paint White $" & paintvalue1 & "<BR>"
response.write "Total Paint Project $" & paintvalue2 & "<BR>"
response.write "<HR>"
response.write "Grand Total = $" & value + value2 + paintvalue1 + paintvalue2
response.write "<HR>"
response.write contvalue 

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

