<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Location Finder to Default Canada to Canada and USA to USA-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Finder</title>
</head>
<body>
<!--#include file="CountryLocation.inc"-->
<%
if CountryLocation = "CANADA" then
	Response.Redirect "index.html"
else
	Response.Redirect "indexTexas.html"
end if
%>
</body>
</html>