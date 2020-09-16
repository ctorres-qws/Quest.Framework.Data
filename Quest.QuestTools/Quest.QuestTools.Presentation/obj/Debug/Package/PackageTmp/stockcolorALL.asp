<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Stock Colour All page gives a large list of all current inventory in all colours organized by part-->
<!--Created for special request by Mary Darnell April 2017-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>All Stock in All colours</title>
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
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE (WAREHOUSE ='GOREWAY' or WAREHOUSE ='HORNER'  or WAREHOUSE ='NASHUA') AND COLOUR <> 'Mill' ORDER BY COlour ASC, PART ASC, Qty DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Pending Stock" selected="true">
         <li class="group">Ordered By Colour, Part, Qty</li>
		 <li><table><TR><TH>COLOUR</TH><TH>PART</TH><TH>QTY</TH><TH>Length</TH></tr>
        
<% 
OldColour = ""
OldPart = ""
OldLength = ""
ShowQTY = 0
do while not rs.eof

OldColour = NewColour
OldPart = NewPart
OldLength = NewLength
NewColour = rs("colour")
NewPart = rs("part")
NewLength = rs("Lft")

if OldColour = NewColour AND OldPart = NewPart AND OldLength = NewLength then

ShowQty = ShowQty + rs("QTY")

else
response.write "<tr>"
response.write "<td>" & OldColour & "</td>"
response.write "<td>" & OldPart & "</td>"
response.write "<td>" & SHowQTY & "</td>"
response.write "<td>" & OldLength & "</td>"
response.write "</tr>"

ShowQty = rs("QTY")

end if
rs.movenext
loop

response.write "<tr>"
response.write "<td>" & NewColour & "</td>"
response.write "<td>" & NewPart & "</td>"
response.write "<td>" & SHowQTY & "</td>"
response.write "<td>" & NewLength & "</td>"
response.write "</tr>"

%>

</table></li>
<li>//END//</li>
/
      </ul>    

<%
rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing
%>



      </ul>                 
            
            
            
       
            
              
               
                
             
               
</body>
</html>
