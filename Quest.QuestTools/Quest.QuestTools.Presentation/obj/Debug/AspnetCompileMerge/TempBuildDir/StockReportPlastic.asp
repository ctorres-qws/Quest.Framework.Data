<!--#include file="dbpath.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!-- Limited to Extrusion Only, Seperate report for Plastic - Michael Bernholtz, at request of Shaun Levy December 2014-->
 <script src="sorttable.js"></script>
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'NPREP' order by PART ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order BY ID DESC"
'Get a Record Set
    'Set RS2 = DBConnection.Execute(SQL2)
    Set RS2 = GetDisconnectedRS(SQL2, DBConnection)
'Aluminum price
'alumprice = 3.95

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->

<table border='1' class='sortable'>
  <tr><th>Part</th>
    <th>Qty</th>
    <th>Length (ft)</th>
    <th>Colour</th>
	<th>Bundle</th>
    <th>LBF</th>
    <th>Value</th>
    </tr>

<%

rs.movefirst
do while not rs.eof
invpart = rs("part")
partqty = rs("Qty")
lmm = CDBL(rs("lmm"))
lft = CDBL(rs("lft"))
linch = CDBL(rs("linch"))
colour = rs("colour")
project = rs("project")
bundle = rs("bundle")
errormsg = 1
NotPl = 1

rs2.filter = "inventoryType = 'Plastic'"
rs2.movefirst
do while not rs2.eof
	if rs2("part") = invpart then 
		lbf = rs2("lbf")
		errormsg = 0
		NotPl = 0
		else
		end if
rs2.movenext
loop
if errormsg = 1 and NotPl =0 then
response.write invpart & " not in inventory master - " & rs("id") & " <BR>"
end if
if NotPl =  0 then
'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

	if lbf > 5 then 
	pricebar = partqty * lbf
	value2 = value2 + pricebar
	else
			tempvalue = (partqty * (lbf * lft ))
	value = value + tempvalue
	end if

'	if colour = "White" Then
'		paintlft = (linch/12) * partqty
'		totalpaintlft = totalpaintlft + paintlft
'	else
'	
'		If colour = "Mill" then
'		else
'			paintlft2 = (linch/12) * partqty
'			totalpaintlft2 = totalpaintlft2 + paintlft2
'		end if
'	
'	end if

%>

  <tr><td><% response.write invpart %></td>
    <td><% response.write partqty %></td>
    <td><% response.write lft %></td>
    <td><% response.write colour %></td>
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
END IF

	end if ' End NotPl code

rs.movenext
loop

%> </Table> <%
'Plastics do not have a colour - Shaun December 2014
'paintvalue1 = (totalpaintlft *0.14)
'paintvalue2 = (totalpaintlft2 *0.38)

response.write "<BR><HR>"
response.write "Total lbf Material $" & round(value,2) & "<BR>"
response.write "Total perbar Material $" & round(value2,2)  & "<BR>"
response.write "<HR>"
'response.write "SubTotal = $" & round(value,2) + round(value2,2) & "<BR><BR>"
'response.write "Total Paint White $" & paintv & "<BR>"
'response.write "Total Paint Project $" & paintvalue2 & "<BR>"
'response.write "<HR>"
response.write "Grand Total = $" & round(value,2) + round(value2,2) 
'+ round(paintvalue1,2) + round(paintvalue2,2)

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

