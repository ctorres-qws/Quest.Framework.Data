<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created Janaury 28th, 2015 - by Michael Bernholtz --> 
<!--Add QT Confirmation page - from GlassOptimaQT.asp Requested by Sasha-->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Edited </title>
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
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassOptimaCardinalSelect.asp" target="_self">Cardinal</a>
    </div>
    
      
    
<form id="conf" title="Glass Edited" class="panel" name="conf" action= "GlassOptimaCardinalSelect.asp" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%       
 
CSent = REQUEST.QueryString("CSent")
CExpected = REQUEST.QueryString("CExpected")
CReceived = REQUEST.QueryString("CReceived")   

	for each item in Request.QueryString("GID")
	GID = item
	GIDList = GIDList & GID & ", " 
		if CExpected <> "" and isNull(CExpected) = False then
			'Set Glass Inventory Update Statement
				StrSQL = "UPDATE Z_GLASSDB  SET [CARDINALExpected]= '" & MONTH(CExpected)&"/"&DAY(CExpected)&"/"&YEAR(CExpected) & "'  WHERE ID = " & GID
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		end if		
		if CReceived <> "" and isNull(CReceived) = False then
			'Set Glass Inventory Update Statement
				StrSQL = "UPDATE Z_GLASSDB  SET [CARDINALReceived]= '" & MONTH(CReceived)&"/"&DAY(CReceived)&"/"&YEAR(CReceived) & "'  WHERE ID = " & GID
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		end if		

	NEXT

%>

    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	

		Response.Write "<li>Cardinal Information Updated:</li>"
		Response.Write "<li>  Updated to : " & GIDList & "</li>"	
		if CExpected <> "" and isNull(CExpected) = False then
			Response.Write "<li> Expected: " & CEXPECTED & "</li>"
		end if		
		if CReceived <> "" and isNull(CReceived) = False then
			Response.Write "<li> RECEIVED: " & CRECEIVED & "</li>"
		end if		
	
		
		

				
%>

        <BR>
       
         <a class="whiteButton" href="javascript:conf.submit()">Optima Select</a>
            
            </form>

            
    
</body>
</html>

<% 

DBConnection.close
set DBConnection=nothing
%>

