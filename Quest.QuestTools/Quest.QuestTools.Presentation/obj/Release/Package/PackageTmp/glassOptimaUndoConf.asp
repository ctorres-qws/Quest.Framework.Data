<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--#include file="dbpath.asp"-->
		<!--Optima Undo Selection Page, resets Optima Date for options chosen in glassOptimaUndo.asp-->
		<!--Created July 2014, at Request of Sasha and Jody to Undo Optima Export-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
</head>

<body>
 <ul id="Profiles" title=" Optima Report" selected="true">
<%

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	Dim item
	For each item in Request.QueryString("OptimaUndo")
		OptimaUndo = item

		Set rs = Server.CreateObject("adodb.recordset")
			strSQL = "UPDATE Z_GLASSDB SET [OPTIMADATE] = NULL, [COMPLETEDDATE] = NULL, [HIDE] = NULL, [SHIPDATE] = NULL  WHERE ID = " & OptimaUndo
			rs.Cursortype = 2
			rs.Locktype = 3
			rs.Open strSQL, DBConnection

		Response.write "<li> Optima Date Removed from ID: " & OptimaUndo
		Response.write "</li>"

	Next

	DbCloseAll

End Function

%>

<a class = 'whiteButton' href = 'GlassOptimaUndo.asp' >Return to Undo </a>

<%

'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>