
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 <!-- February 2019 - Added USA view -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>

<%
part = request.QueryString("part")

id = REQUEST.QueryString("ID")
aisle = REQUEST.QueryString("aisle")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle">All Stock by PO/Bundle</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>

              <form id="edit" title="Select Stock by PO" class="panel" name="edit" action="allbypobundle.asp" method="GET" target="_self" selected="true" > 
        <h2>Search ALL Stock by PO / Bundle</h2>

<fieldset>
     <div class="row">
                <label>Search</label>
                <input type="text" name='pobundle' id='pobundle'>
            </div>
            
</fieldset>

        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Search ALL stock by PO / Bundle</a><BR>
            
            </form> 

</body>
</html>

<%
DBConnection.close
set DBConnection=nothing
%>

