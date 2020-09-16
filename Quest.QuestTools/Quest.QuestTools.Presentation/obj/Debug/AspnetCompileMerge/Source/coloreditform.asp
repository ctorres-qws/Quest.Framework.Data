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

cid = REQUEST.QueryString("CID")

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="coloredit.asp" target="_self">Edit Color</a>
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->




Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_COLOR"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & cid



%>
    
    
              <form id="edit" title="Edit Color" class="panel" name="edit" action="coloreditconf.asp" method="GET" target="_self" selected="true" > 
        <h2>Add Color</h2>
  
        <fieldset>


        <div class="row">
            <label>Job Code</label>
            <input name='Job' type="text" id='Job' value="<% response.write rs.fields("job") %>" >
        </div>
		
		         <div class="row">
            <label>Ext / Int</label>
            <Select name='Side'>
				<option value="" selected>Please Select</option>
			<!--	<option value="<%response.write rs.fields("Side")%>" selected> <%response.write rs.fields("Side")%></option> -->
				<option value="Ext." <% if rs.Fields("Side") = "Ext." then Response.write "Selected"%> >Ext</option>
				<option value="Int." <% if rs.Fields("Side") = "Int." then Response.write "Selected"%> >Int</option>
			</Select> 
		</div>

		<div class="row">
                <label>Paint Code</label>
                <input type="text" name='CODE' id='CODE' value="<% response.write rs.fields("code") %>" >
        </div>
		
		<div class="row">
            <label>Paint Type</label>
            <select name='Company'>
				<option value="<% response.write rs.fields("company") %>"><% response.write rs.fields("company") %></option>
				<option value="PPG Acrynar">PPG Acrynar</option>
				<option value="PPG Duranar">PPG Duranar</option>
				<option value="PPG Duracron">PPG Duracron</option>
				<option value="PPG Duracron White">PPG Duracron White K-1285</option>
				<option value="PPG Duranar XL">PPG Duranar XL</option>
				<option value="PPG Duranar XL">PPG Duranar XL + Basecoat</option>
				
				<option value="VALSPAR Acrodize">VALSPAR Acrodize</option>
				<option value="VALSPAR Acroflur">VALSPAR Acroflur</option>
				<option value="VALSPAR Clear Anodize">VALSPAR Clear Anodize</option>
				<option value="VALSPAR Fluropon">VALSPAR Fluropon</option>
				<option value="VALSPAR Fluropon Classic">VALSPAR FluroponcClassic</option>
				<option value="VALSPAR Flurospar">VALSPAR Flurospar</option>
				<option value="VALSPAR Polylure">VALSPAR Polylure</option>

				<option value="Other">Other</option>
			</Select>
        </div>
		
        <div class="row">
            <label>Description</label>
            <input type="text" name='DESCRIPTION' id='DESCRIPTION' value="<% response.write rs.fields("desc") %>" >
        </div>

            
        <div class="row">
            <label>Price Cat.</label>
            <input type="text" name='PAINTCAT' id='PAINTCAT' value="<% response.write rs.fields("pricecat") %>" >
        </div>
		
		<div class="row">
            <label>Active</label>
            <input type="checkbox" name='Active' id='Active' <% if rs.fields("ACTIVE") = TRUE THEN response.write "checked" END IF%> >
        </div> 	
		
				<div class="row">
            <label>Extrusion</label>
            <input type="checkbox" name='EXTRUSION' id='EXTRUSION' <% if rs.fields("EXTRUSION") = TRUE THEN response.write "checked" END IF%>>
        </div> 	
		
		<div class="row">
            <label>Sheet</label>
            <input type="checkbox" name='SHEET' id='SHEET' <% if rs.fields("SHEET") = TRUE THEN response.write "checked" END IF%> >
        </div> 	
		
		
            <input type="hidden" name='cid' id='cid' value="<%response.write rs.fields("id") %>" >
                      
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
            
            
      
            
            </form> 

  
<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

          
    
</body>
</html>