<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 

<html xmlns="http://www.w3.org/1999/xhtml">
<head>

</head>
<body>
  	 <!--#include file="dbpath.asp"-->
<%
response.write "<p>Hello</p>"

ReturnSite = request.querystring("ReturnSite")
bocid = request.querystring("bocid")
strSQL = FixSQL("UPDATE X_BACKORDER SET ACTIVE = FALSE , ReorderDate = #" & Now & "# WHERE ID = " & bocid)
Set RSComplete = DBConnection.Execute(strSQL)
DBConnection.close
set DBConnection=nothing

Response.Redirect Returnsite
%>
</body>
</html>