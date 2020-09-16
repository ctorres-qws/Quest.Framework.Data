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
  
<% 

cid = REQUEST.QueryString("CID")

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PanelStylebyJob1.asp" target="_self">Panel Style</a>
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->




Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM StylesPanel"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & cid



%>
    
    
              <form id="edit" title="Edit Panel Style" class="panel" name="edit" action="PanelStyleeditconf.asp" method="GET" target="_self" selected="true" > 
        <h2>Edit Panel Style</h2>
  
        <fieldset>


        <div class="row">
            <label>Parent Job</label>
            <input name='Parent' type="text" id='Parent' value="<% response.write rs.fields("Parent") %>" >
        </div>
		
		<div class="row">
            <label>Colour Code</label>
            <input name='ColorCode' type="text" id='ColorCode' value="<% response.write rs.fields("ColorCode") %>" >
        </div>
		
		<div class="row">
            <label>Style Name</label>
            <input name='Name' type="text" id='Name' value="<% response.write rs.fields("Name") %>" >
        </div>
		
		<div class="row">
                <label>Description</label>
                <input type="text" name='Description' id='Description' value="<% response.write rs.fields("Description") %>" >
        </div>
		
		         <div class="row">
            <label>Ext / Int</label>
            <Select name='Side'>
				<option value="<%response.write rs.fields("Side")%>" selected> <%response.write rs.fields("Side")%></option>
				<option value="Ext.">Exterior</option>
				<option value="Int.">Interior</option>
			</Select> 
		</div>


		
		<div class="row">
            <label>Material</label>
            <select name='Material'>
				<option value="<% response.write rs.fields("Material") %>"><% response.write rs.fields("Material") %></option>
					<option value="0.050 INCH ALUM">0.050'' ALUM</option>
					<option value="0.080 INCH ALUM">0.080'' ALUM</option>
					<option value="0.125 INCH ALUM">0.125'' ALUM</option>
					<option value="Steel">Steel</option>

			</Select>
        </div>
		
        <div class="row">
            <label>Colour</label>
            <input type="text" name='Colour' id='Colour' value="<% response.write rs.fields("Colour") %>" >
        </div>
		
		<div class="row">
            <label>Notes</label>
            <input type="text" name='Notes' id='Notes' value="<% response.write rs.fields("Notes") %>" >
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