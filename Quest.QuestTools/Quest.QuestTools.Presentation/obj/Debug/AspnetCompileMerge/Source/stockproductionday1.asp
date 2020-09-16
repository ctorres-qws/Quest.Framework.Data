<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Stock Entered Today and Production Today Date Choice added for Alex Sofienko -->
<!-- October 2014, Michael Bernholtz-->
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
currentDate = Date


%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self" >Inventory</a>
    </div>
    
        
    
              <form id="CDate" title="Date for Stock/Prod" class="panel" name="CDate"  method="GET" target="_self" selected="true" > 
			  <h2>Choose a Date for Stock Entry or Sent to Production (DD/MM/YYYY)</h2>

<fieldset>
     <div class="row">
	 <%
	 InputDate = Date()
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
	 %>
                <label>Day</label>
                <input type="date" name='CDay' id='CDay' value = '<%Response.write DateInput%>' /> 
            </div>
            
</fieldset>


        <BR>

        <a class="greenButton" onClick="CDate.action='StockToday.asp'; CDate.submit()">Stock Entered</a><BR>
		<a class="greenButton" onClick="CDate.action='ProductionToday.asp'; CDate.submit()">Sent to Production</a><BR>
          
            </form> 
            
</html>


