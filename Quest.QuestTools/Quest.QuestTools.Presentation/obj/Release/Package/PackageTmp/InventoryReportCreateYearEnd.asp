<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Created at Request of Shaun Levy with permission from Jody Cash -->
<!--Input form to take snapshot of Y_Inv and add to SQL DATABASE using new DBPATH in next page-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Inventory Snapshot</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%

cMonth = month(now)
cYear = year(now)

if cMonth = 1 then 
	cMonthy =  12
	cYeary = cYear-1
else
	cMonthy = cMonth-1
	cYeary = cYear
end if 

if cMonth < 10 then
	cMonth = "0" & cMonth
end if

if cMonthy < 10 then
	cMonthy = "0" & cMonthy
end if
 %>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
    <form id="snapshot" title="Take Inventory Snapshot" class="panel" name="snapshot" action="InventoryReportConfYearEnd.asp" method="GET" target="_self" selected="true">
        <h2>Select Date for Snapshot</h2>
        <fieldset>

			<div class="row">
			<label>Month</label>
			<select id='snapmonth' name="snapmonth" >
				<option value="previous"><% response.write cMonthy & "/" & CYeary %></option>
				<option value="current"><% response.write cMonth & "/" & CYear %></option>

			</select>
            </div>

		</fieldset>
        <BR>
        <a class="whiteButton" href="javascript:snapshot.submit()">Submit</a>  

            </form>

</body>
</html>
