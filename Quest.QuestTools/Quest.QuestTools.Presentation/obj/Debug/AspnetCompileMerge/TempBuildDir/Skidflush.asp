<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Flush - Sets Flush to True and Adds a Flush date -->
<!-- Can be autocalled by scanadd.asp with an add flag-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Scan to Skid</title>
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

<body>
     
     <% 
	 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select distinct name FROM SKIDItem order by name ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection 

	 
	 
	 
currentDate = Date()

IsError = False
' Reset the Variable for locating an Error
	 
skid = UCASE(request.querystring("skidname"))
add = UCASE(request.querystring("add"))
			
			SQL1 = "UPDATE SkidItem Set Flushed = true, FlushedDate ='" & currentDate & "' WHERE name = '" & skid & "' AND Flushed = false"
			Set RS1 = DBCOnnection.Execute(SQL1)
	

 %>

     


    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
   
   <form id="flushskid" title="Flush Skid" class="panel" name="flushskid" action="skidflush.asp" target="_self" method="GET" selected="true">
         <% if skid = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write skid & " - Flushed <BR>" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Empty the Skid</h2>
        <fieldset>
       
            
                <div class="row">
                <label>Skid</label>
					<select name='skidname' id='skidname'>
						<%
						rs.movefirst
						do while not rs.eof
							Response.Write "<option value = '"
							Response.Write TRIM(rs("name"))
							Response.Write "'>"
							Response.Write TRIM(rs("name"))
							Response.Write "</option>"
						rs.movenext
						loop
						%>

					</select>
               <!-- </div>-->
				</div>
            
                  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("skidname");
  
        textbox.value = barcode;
    
}
  

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
            <input type="submit" class="redButton" onsubmit=="return confirm('Are you sure you want to do that?');">
	<%		
		 if add = "1" then
			response.write" <a class='whiteButton' target='#_self' href='skidadd.asp'>Back to Add</a>"
		end if
	%>

			</form>
<%
rs.close
set rs = nothing

DBConnection.close
set DBConnection = nothing
%>
</body>
</html>