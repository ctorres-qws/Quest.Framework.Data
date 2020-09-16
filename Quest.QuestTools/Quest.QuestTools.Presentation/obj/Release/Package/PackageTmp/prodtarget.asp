<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth, initial-scale=1.0, maximum-scale=1.0, user-scalable=0"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  
  


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

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
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html" target="_self">Home</a>
				<!--#include File="inc_style.asp"-->
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
ID = REQUEST.QueryString("ID")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM X_PRODTARGETS"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection



%>
    
    
              <form id="edit" title="Edit Targets" class="panel" name="edit" action="prodtargetconf.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Targets" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Production Targets</h2>
  
   

<fieldset>
     <div class="row">
                <label>Glazing D</label>
                <input type="text" name='DAYG' id='DAYG' value="<%response.write rs.fields("Target") %>">
            </div>
<%rs.movenext%>
                        <div class="row">
                <label>Glazing N</label>
                <input type="text" name='NIGHTG' id='NIGHTG' value="<%response.write rs.fields("Target") %>">
            </div>
            <%rs.movenext%>
         <div class="row">

                        <div class="row">
                <label>Assembly D</label>
                <input type="text" name='DAYA' id='DAYA' value="<%response.write rs.fields("Target") %>">
            </div>
            <%rs.movenext%>
              <div class="row">
                <label>Assembly N</label>
                <input type="text" name='NIGHTA' id='NIGHTA' value="<%response.write rs.fields("Target") %>">
            </div>
            
            <%rs.movenext%>
            
                     <div class="row">

                        <div class="row">
                <label>IG D</label>
                <input type="text" name='DAYIG' id='DAYIG' value="<%response.write rs.fields("Target") %>" >
            </div>
            <%rs.movenext%>
                           <div class="row">
                <label>IG N</label>
                <input type="text" name='NIGHTIG' id='NIGHTIG' value="<%response.write rs.fields("Target") %>" >
            </div>
            
               
                  
            
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
            
            <input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>">

            
            </form> </form>
            
 <form id="conf" title="Edit Stock" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Stock Edited</h2>
  

            
            </form>
            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set conntemp=nothing
%>

