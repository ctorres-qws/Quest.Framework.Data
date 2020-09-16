<!--#include file="dbpath.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT yI.* FROM Y_INV yI WHERE (yI.WAREHOUSE IN('GOREWAY','NASHUA','DURAPAINT','HORNER')) AND (yI.Part = 'Que-162' OR yI.Part = 'Que-163' OR yI.Part = 'Que-164' OR yI.Part = 'Que-165' OR yI.Part = 'Que-166' OR yI.Part = 'Que-167' OR yI.Part = 'Que-172') order by PART ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
'rs.filter = "Part = 'Que-162' OR Part = 'Que-163' OR Part = 'Que-164' OR Part = 'Que-165' OR Part = 'Que-166' OR Part = 'Que-167' OR Part = 'Que-172'"

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
'Get a Record Set

Set rs2 = GetDisconnectedRS(SQL2, DBConnection)
'Aluminum price
alumprice = 3.95

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<H2> Prep Material </H2>
<H3>Que 162-167, Que 172 - in Goreway, Horner, and Durapaint</h3>
<table border='1' class='sortable'>
  <tr><th>Part</th>
    <th>Qty</th>
    <th>Length (mm)</th>
    <th>Colour</th>
	 <th>Bundle</th>
   <!-- <th></th>-->
    <th>Value</th>
    </tr>

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

rs2.movefirst
do while not rs2.eof
if UCase(rs2("part")) = UCase(invpart) then 
kgm = rs2("kgm")
errormsg = 0
'Response.write rs2("kgm") & "<BR>"
else
end if
rs2.movenext
loop

if errormsg = 1 then
response.write invpart & " not in inventory master - " & rs("id") & " <BR>"
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
	<td><% response.write bundle %></td>
  <!--  <td></td> -->
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
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

