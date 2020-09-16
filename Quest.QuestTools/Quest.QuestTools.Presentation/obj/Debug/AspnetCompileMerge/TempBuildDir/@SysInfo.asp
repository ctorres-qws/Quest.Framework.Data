<!--#include file="@common.asp"-->
<%

Response.Write("FolderUploadRecords: " & gstr_FolderUploadRecords & "<br/>")

Response.Write("SQL_Server: " & b_SQL_Server & "<br/>")

Response.Write("Your IP: " & Request.ServerVariables("REMOTE_ADDR") & "<br/>")

Set DBConn = Server.CreateObject("adodb.connection")

DSN = GetConnectionStr(true) 'method in @common.asp
DBConn.Open DSN
DBConn.Close

Response.Write("<br />")

Response.Write("Database Opened: " & DSN)

'Response.Write("<br/>Test:" & FixSQLCheck("SELECT top 1000 * FROM X_Shipping_Truck WHERE active = TRUE {0} ORDER BY ID DESC", false))
isSQLServer = false
Response.Write("<br/>" & FixSQLCheck("SELECT top 1000 * FROM X_Shipping_Truck WHERE active = TRUE {0} ORDER BY ID DESC", isSQLServer))




%>