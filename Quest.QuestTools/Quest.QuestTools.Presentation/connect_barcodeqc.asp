<!--#include file="dbpath.asp"-->
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_EMPLOYEES"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

%>