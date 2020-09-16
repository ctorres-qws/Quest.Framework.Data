<!--#include file="dbpath.asp"-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
<!-- Hardware Report based on Inventory Stock Report - Requested by Shaun levy and Kevin Cosgrave March 2018 -->
 <script src="sorttable.js"></script>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_HARDWARE WHERE WAREHOUSE = 'NASHUA' order by PART ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

'Create a Query
    SQL2 = "SELECT * FROM Y_HARDWARE_MASTER order BY ID DESC"
'Get a Record Set
'Set RS2 = DBConnection.Execute(SQL2)
Set RS2 = GetDisconnectedRS(SQL2, DBConnection)


'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<P> Reporting Value of HARDWARE Inventory in Nashua</P>

<table border='1' class='sortable'>
	<tr>
	<th>Part</th>
	<th>Description</th>
    <th>Qty</th>
	<th>PO</th>
	<th>EnterDate</th>
	<th>LastModify</th>
    <th>Price Per Unit</th>
    <th>Value</th>
    </tr>
<%
Server.ScriptTimeout=400
%>
<%

rs.movefirst
do while not rs.eof
	part = rs("part")
	qty = rs("Qty")
	po = rs("po")
	EnterDate = rs("EnterDate")
	LastModify = rs("LastModify")
	errormsg = 1
	NotEx = 1


	rs2.movefirst
	do while not rs2.eof
		if rs2("part") = part then 
			Price = RS2("Price")
			Description = RS2("Description")
			errormsg = 0
			NotEx = 0
	Exit Do
		else
		end if
	rs2.movenext
	loop
	if errormsg = 1 and NotEx =0 then
	response.write part & " not in Hardware inventory master - " & rs("Part") & " <BR>"
	end if
	if NotEx =  0 then

	If Not IsNull(QTY) Then

	TempValue = qty * Price
	TotalValue = TotalValue + TempValue	

%>

	<tr>
	<td><% response.write part %></td>
    <td><% response.write description %></td>
    <td><% response.write qty %></td>
    <td><% response.write po %></td>
	<td><% response.write EnterDate %></td>
	<td><% response.write LastModify %></td>
	<td><% response.write Price %></td>
    <td>$<% response.write round(tempvalue,2) %></td>
    </tr>

<%	

ELSE
RESPONSE.WRITE "QTY IN INVENTORY IS ZERO <BR>"
END IF

	end if ' End NotEx code

rs.movenext
loop

%> </Table> <%


response.write "<BR><HR>"
response.write "Total Hardware Value $" & round(Totalvalue,2) & "<BR>"
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