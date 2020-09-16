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
	strSQL = "SELECT * FROM Y_HARDWARE WHERE [MOVEDBY] = 'MW' ORDER BY WAREHOUSE, PART"
Else
	strSQL = "SELECT * FROM Y_HARDWARE WHERE [MOVEDBY] = 'MW' ORDER BY WAREHOUSE, PART"
End If
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_HARDWARE_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

%>
 
    </head>
<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=STATUSCOUNTHARDWARE_2018.xls"
%>
<body>
        <ul id="Profiles" title="Stock WIth Status Notes" selected="true">
         
<% 
if not rs.eof then
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><TH>Warehouse</TH><th>Part</th><th>Description</th><th>New Quantity</th><th>Old Qty</TH><TH>Delta</th><th>PO</th><th>Aisle</th><th>Rack</th><th>Level</th><th>Status</th><th>New Value</th><th>Old Value</th><th>Value Delta</th></tr>"
end if

do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
		Price = 0
		
	else
		Description = rs2("Description")
		Price = rs2("Price")
	end if
	
'	Set rs3 = Server.CreateObject("adodb.recordset")
'	strSQL3 = "SELECT * FROM Y_Hardware_Log WHERE HARDWAREID = " & rs("ID") & " AND TRANSACTION = 'original' ORDER BY PART ASC"
'	rs3.Cursortype = 2
'	rs3.Locktype = 3
'	rs3.Open strSQL3, DBConnection
'	
'	
'	if rs3.EOF or rs3.BOF THEN
'		OLDQTY = 0
'	else
'	
'		OLDQTY = 0
'		OLDQTY = rs3("QTY")
'	end if
	
'	rs3.close
'	set rs3 = nothing
	
	Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT * FROM Y_HARDWARE_LOG WHERE HardwareID = " & rs("ID") & " ORDER BY PART ASC, ID ASC"
	rs3.Cursortype = 2
	rs3.Locktype = 3
	rs3.Open strSQL3, DBConnection
	rs3.filter = "[MOVEDBY] = 'MW' and TRANSACTION = 'edit'"
	
	if rs3.eof then
	
	OLDQTY = rs("QTY")
	
	else
	
		ChangeID = CDBL(rs3("ID")) - 1
		rs3.filter = "[ID] = " & ChangeID
		
		OLDQTY = 0
		OLDQTY = RS3("QTY")
	
	end if
	
	rs3.close
	set rs3 = nothing
	
	

Response.write "<tr>"
Response.write "<td>" & rs("WAREHOUSE") & "</td>"
Response.write "<td>" & rs("part") & "</td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & OLDQTY & " </td>"
Response.write "<td>" & OLDQTY - rs("Qty") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Aisle") & " </td>"
Response.write "<td>" & rs("Rack") & " </td>"
Response.write "<td>" & rs("Level") & " </td>"
Response.write "<td>" & rs("MOVEDBY") & " </td>"
Response.write "<td>$" & Round(rs("qty") * Price,2) & " </td>"
Response.write "<td>$" & Round(OLDQTY * Price,2) & " </td>"

if OldQTY >= rs("QTY") then
	Response.write "<td>$" & Round((OLDQTY - rs("Qty"))* Price,2) & " </td>"
else
	Response.write "<td><font color = 'red'>$" & Round((OLDQTY - rs("Qty"))* Price,2) & " </FONT></td>"
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
