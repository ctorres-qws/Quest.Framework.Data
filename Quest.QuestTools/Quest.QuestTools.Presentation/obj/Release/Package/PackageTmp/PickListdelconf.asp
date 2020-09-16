<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

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
strSQL = "SELECT * FROM PICKLIST ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

PKid = REQUEST.QueryString("PKID")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PickListEdit.asp%>" target="_self">Manage PL</a>
    </div>
    
      <%                  
PKid = request.querystring("PKid")

'Set Pick List Delete Statement
				StrSQL = "DELETE FROM PICKLIST WHERE ID = " & PKid
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	

DBConnection.close
set DBConnection=nothing	 

%>
		<form id="conf" title="Delete" class="panel" name="conf" action="PickListEdit.asp" method="GET" target="_self" selected="true" >   
        <h2>Pick List Item: Removed</h2>



        <BR>
        
		<a class="whiteButton" href="javascript:conf.submit()">Back to Pick List Manager</a>
            
            </form>
 
<% 


%>

           
    
</body>
</html>

