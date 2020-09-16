<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created Janaury 28th, 2015 - by Michael Bernholtz --> 
<!--Add QT Confirmation page - from GlassOptimaQT.asp Requested by Sasha-->
<!--Based on Comma and entered ids instead of selection -->
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
                <a class="button leftButton" type="cancel" href="GlassOptimaQTSelect.asp" target="_self">Add QT</a>
    </div>
    
      
    
<form id="conf" title="Glass Edited" class="panel" name="conf" action= "GlassOptimaQTSelect2.asp" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%       
 
QT = REQUEST.QueryString("QT")
IDLIST = Request.QueryString("IDList")
IDLIST = IDLIST & ","
IDLIST = Replace(IDlist," ","")
GIDLIST = IDLIST
counter =1

DO while INSTR(1, IDLIST, ",",1)
counter =counter+1
CommaPlace = Instr(1, IDLIST, ",",1)
GID = LEFT(IDLIST,CommaPlace-1) 

IDLIST = Right(IDLIST, LEN(IDLIST)- CommaPlace)

'Set Glass Inventory Update Statement
				StrSQL = "UPDATE Z_GLASSDB  SET [QTFILE]= '" & QT & "', [QuickTempSent]= '" & MONTH(DATE) & "/" & DAY(DATE) & "/" &YEAR(DATE) & "'  WHERE ID = " & GID
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
				
loop


%>

    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
		Response.Write "<li>QT ADDED:</li>"
		Response.Write "<li> QTs Added to : " & GIDList & "</li>"
		Response.Write "<li> QTs: " & QT & "</li>"

				
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

