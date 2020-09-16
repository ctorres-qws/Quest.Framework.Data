<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search Production Stock by PO Search page -->
<!--Created May 1st, by Michael Bernholtz at Request of Ruslan Bedoev -->
<!-- February 2019 - Added USA view -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Production Inventory</title>
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
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Production by PO/Bundle</h1>
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
    
        
    
              <form id="edit" title="Select Stock by PO" class="panel" name="edit" action="productionbypo.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Stock" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Search Production Stock by PO</h2>
  

<fieldset>
     <div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO'>
            </div>
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Search Production stock by PO</a><BR>
            
            </form> </form>
            
 <form id="conf" title="Edit Stock" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Stock Edited</h2>
  

            
            </form>
<%
DBConnection.close
Set DBConnection = nothing
%>
Set 	
</body>
</html>


