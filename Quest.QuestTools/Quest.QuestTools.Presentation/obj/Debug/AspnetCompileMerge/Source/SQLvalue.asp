<!--#include file="dbpath.asp"-->
<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

counter = 0

rs.movefirst
do while not rs.eof
part = rs("Part")
response.write part
counter = counter + 1

if counter = 100 OR counter = 200 or counter = 300 or counter = 400 or counter = 500 or counter = 600 then
Response.flush
end if
rs.movenext
loop

response.write "done"
%>

       
            
    
</body>
</html>

<% 

rs.close
set rs=nothing
rs2.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

