<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stocktoday.asp duplicated and put into table form, at Request of Ruslan Bedoev, May 23rd, 2014-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>

    </script>
    
    <%
	currentDate = Request.Querystring("CDay")
	CDay = currentDate  
	if currentDate = "" then
		currentDate = Date()
	End if
	
Set rs = Server.CreateObject("adodb.recordset")
If b_SQL_Server Then
	strSQL = "SELECT * FROM Y_INV WHERE [NOTE 2] = 'MW' ORDER BY WAREHOUSE, PART"
Else
	strSQL = "SELECT * FROM Y_INV WHERE [NOTE 2] = 'MW' ORDER BY WAREHOUSE, PART"
End If
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

%>
 
    </head>
<%
'Response.ContentType = "application/vnd.ms-excel"
'Response.AddHeader "Content-Disposition", "attachment; filename=STATUSCOUNT_2018.xls"
%>
<body>

        <ul id="Profiles" title="Stock WIth Status Notes" selected="true">
         
<% 
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><TH>Warehouse</TH><th> New Quantity</th><TH>Old Qty</TH><TH>DELTA</TH><th>Colour</th><th>PO</th><th>Bundle</th><th>Length(Ft)</th><th>Status</th><th>Extr Value</th><TH>Old Extr Value</TH><TH>DELTA</TH></tr>"

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
		KGM = 0
		Price = 0
		
	else
		Description = rs2("Description")
		KGM = rs2("KGM")
		Price = KGM 
	end if
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT * FROM Y_INVLOG WHERE ITEMID = '" & rs("ID") & "' ORDER BY PART ASC, ID ASC"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL3, DBConnection
	rs3.filter = "[NOTE 2] = 'MW' and TRANSACTION = 'edit'"
	if rs3.eof then
		OLDQTY = 0
	else
		ChangeID = rs3("ID")-1
		rs3.filter = "[ID] = " & ChangeID
		
		OLDQTY = 0
		OLDQTY = RS3("QTY")
	end if
	rs3.close
	set rs3 = nothing

Response.write "<tr>"
Response.write "<td>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & RS("Warehouse") & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & OLDQTY & " </td>"
Response.write "<td>" & OLDQTY - rs("Qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Bundle") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note 2") & " </td>"

if PRICE > 5 then
	Response.write "<td>" & Round(Price * RS("QTY") * 3.95,2) & " </td>"
	Response.write "<td>" & Round(Price * OLDQTY * 3.95 ,2) & " </td>"
	if OldQTY >= rs("QTY") then
		Response.write "<td>" & Round(Price * (OLDQTY - RS("QTY")),2) & " </td>"
	else
		Response.write "<td><font color = 'red'>" & Round(Price * (OLDQTY - RS("QTY")),2) & " </font></td>"
	end if
else
	Response.write "<td>" & Round(Price * (CDBL(RS("LMM")) / 1000) * RS("QTY") * 3.95,2) & " </td>"
	Response.write "<td>" & Round(Price * (CDBL(RS("LMM")) / 1000) * OLDQTY * 3.95 ,2) & " </td>"
	if OldQTY >= rs("QTY") then
		Response.write "<td>" & Round(Price * (CDBL(RS("LMM")) / 1000) * (OLDQTY - RS("QTY")),2) & " </td>"
	else
		Response.write "<td><font color = 'red'>" & Round(Price * (CDBL(RS("LMM")) / 1000) * (OLDQTY - RS("QTY")),2) & " </font></td>"
	end if
end if



Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"




rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>                 
            
            
            
       
            
              
               
                
             
               
</body>
</html>
