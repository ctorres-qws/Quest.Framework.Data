<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
'
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_BARCODE_LINEITEM ORDER BY JOB ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

g2id = REQUEST.QueryString("g2ID")
rs.filter = " ID = " & g2id


%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="glazing2report.asp" target="_self">Glazing 2</a>
    </div>
    
    
    <form id="edit" title="Add Glazing 2 Reason" class="panel" name="edit" action="glazing2reportconf.asp" method="GET" target="_self" selected="true" > 
	<h2>Glazing 2 Reason </h2>
          </h2>

<fieldset>
            <div class="row">
                <label>Job: <% response.write rs("job") %></label>
            </div>
			<div class="row">
                <label>Floor: <% response.write rs("floor") %></label>
            </div>
			<div class="row">
                <label>Reason</label>
                <input type="text" name='g2reason' id='g2reason' value="<%response.write rs.fields("g2reason")%>" >
            </div>
			
                      
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
            
            <input type="hidden" name='g2id' id='g2id' value="<%response.write g2id%>">
      
            
            </form> 
            

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

