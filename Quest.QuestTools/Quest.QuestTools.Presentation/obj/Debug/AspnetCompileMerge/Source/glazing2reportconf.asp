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

	</head>
<body  >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="glazing2report.asp" target="_self">Glazing 2</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="glazing2report.asp" method="GET" target="_self" selected="true" >              

  
   
        <h2>SGlazing 2 Reason Added</h2>
  
<%                  
g2id = request.querystring("g2id")


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_BARCODE_LINEITEM"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & g2id


g2reason = request.querystring("g2reason")
rs.Fields("g2reason") = g2reason
rs.update
%>


        <BR>
       
        <ul>
		<li><% response.write rs("job") & " - " & rs("floor")%></li>
		<li><% response.write g2reason%></li>

		</ul>
         <a class="whiteButton" href="javascript:conf.submit()">Back to Glazing 2</a>
            
            </form>

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

