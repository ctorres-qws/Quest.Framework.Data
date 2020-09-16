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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>


<% 
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM PICKLIST ORDER BY JOB ASC, FLOOR ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="index.html#_Pick" target="_self" >Pick List</a>
    </div>
    
        
    
              <form id="edit" title="Select Stock by PO" class="panel" name="edit" action="PickListView.asp" method="GET" target="_self" selected="true" > 
        <h2>Choose the Job/Floor combinations to View</h2>
  
<fieldset>

<%
' Design the list of all search items by finding Unmatching Job and Floor records
if rs.eof then
rs.movefirst
end if


do while not rs.eof
if JOB1 = "" and FLOOR1 = "" then
'Removes blank first record
else
if rs("JOB") = JOB1 and rs("FLOOR") = FLOOR1 then
'Ensures not to Duplicate Job/Floor records
else
%>
<input type="checkbox" name="JobFloor" value="<% response.write JOB1 & "-" & FLOOR1 %>"><% response.write JOB1 & FLOOR1 %><br>
<%
end if
end if
JOB1 = rs("JOB")
FLOOR1 = rs("FLOOR")

rs.movenext
loop
%>
<input type="checkbox" name="JobFloor" value="<% response.write JOB1 & "-" & FLOOR1 %>"><% response.write JOB1 & FLOOR1 %><br>
<%
%>
    

      </fieldset> <BR>
        <a class="lightblueButton" href="javascript:edit.submit()">Search Pick List by Die</a><BR>
	
		
		
            
            </form> 
            

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

