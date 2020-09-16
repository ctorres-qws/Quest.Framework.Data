  <!--#include file="dbpath.asp"-->                     
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--Optima Selection Page for adding Cardinal Data information, shows all items that do not have an Completed Date and a checkbox-->
		<!--Created January 2015, at Request of Sasha for adding a note to multiple items at once-->
		<!-- Sends to glassOptimaCardinalConf.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    <style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>
    <%
	
AddSent = Request.form("CSent")
AddExpected = Request.form("CExpected")
AddReceived = Request.form("CReceived")
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection




%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       <form id="Optima" action="glassOptimaCardinalConf.asp" name="Optima"  method="GET" target="_self" selected="true" >  
        
		<h2><center>Update Cardinal in the fields Below, Then Select the items, then press Submit<center></h2>
		
		<fieldset>

		<div class="row">
                <label>Expected Return</label>
                <input type="date" name='CExpected' id='CExpected' />
        </div>
		<div class="row">
                <label>Received</label>
                <input type="date" name='CReceived' id='CReceived' />
        </div>
               <input type="hidden" name='ticket' id='ticket' value = 'multiple' />     
		<a class="whiteButton" onClick="Optima.action='GlassOptimaCardinalConf.asp'; Optima.submit()">Update Cardinal</a><BR>
		</fieldset>
        <ul id="Profiles" title=" Optima Report" selected="true">
		
		
<% 



' and isdate(RS("CARDINALSENT"))  Can only be put in once Cardinal Sent is automatic

response.write "<li class='group'>Choose Records below to add the QT</li>"
response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>Notes</th><th>TimeLine</th></tr>"
do while not rs.eof
if not isdate(RS("COMPLETEDDATE")) then
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td><td>" & RS("NOTES") & "</td>"
	response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=CardinalUpdate' target ='#_blank' >Time Line</a> </td>"

	response.write "<td><input type='checkbox' name='GID' value='" & RS("ID")& "'></td></tr>"
end if	
	rs.movenext
loop



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing



%>

	</table>
	
      </ul>    
		</form>
            
            
            
       
            
              
               
                
             
               
</body>
</html>
