<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--#include file="countrylocation.inc"-->
<!-- Created February 6th, by Michael Bernholtz - Entry Form to input items to QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
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
    <div class="toolbar">
        <h1 id="pageTitle">Add Items to QC Inventory</h1>
        <a class="button leftButton" type="cancel" href="index.html#_QC" target="_self">QC</a>
        </div>

              <form id="enter" title="IP ADDRESS" class="panel" name="enter" action="Michaeltestpage.asp" method="GET" target="_self" selected="true">
              
			  <h2>System Calls</h2>
              <fieldset>               
       
	   <div class="row">
            <label>Remote Addr: <%response.write Request.ServerVariables("REMOTE_ADDR")%></label>
        </div>
	   <div class="row">
            <label>http X forwarded For: <%response.write Request.ServerVariables("HTTP_X_FORWARDED_FOR")%></label>
        </div>
	   <div class="row">
            <label>Remote IP: <%response.write Request.ServerVariables("HTTP_REMOTE_IP")%></label>
        </div>
	   <div class="row">
            <label>HTTP Host: <%response.write Request.ServerVariables("HTTP_HOST")%></label>
        </div>
	   <div class="row">
            <label>Remote Host: <%response.write Request.ServerVariables("REMOTE_HOST")%></label>
        </div>
	   <div class="row">
            <label>CountryLocation Include based on Texas IP Range: <%response.write CountryLocation%></label>
        </div>
	   <div class="row">
            <label><% response.write Now%> </label>
        </div>		
         
         
</fieldset>


            
            </form>
        
             
               
</body>
</html>
