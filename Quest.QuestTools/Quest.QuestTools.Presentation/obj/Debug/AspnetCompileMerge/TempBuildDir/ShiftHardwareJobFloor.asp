                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Job Floor Entry Page For Shift Hardware View - Searches for a Job and Floor and then shows Each Buggy, Bin, Cart, Container-->
<!-- Created November, 2018 by Michael Bernholtz - Requested by Ariel Aziza-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shift Hardware</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_HW" target="_self">Hardware</a>
        </div>
<form id="Shift" title="Shift Hardware" class="panel" name="Shift" action="ShifthardwareView1.asp" method="GET" selected="true" target="_self">
<fieldset>

<%
JOB = Request.Querystring("JOB")
FLOOR = Request.Querystring("FLOOR")
%>
    
         <div class="row">   
            <label>Job </label>
            <input type="text" name='Job' id='Job' value = "<%response.write Job%>" >
		</div>
		<div class="row">     
            <label>Floor </label>
            <input type="text" name='Floor' id='Floor' value = "<%response.write Floor%>">
			
			
			<input type="hidden" name='PositionX' id='Floor' value = "0">
			<input type="hidden" name='PositionY' id='Floor' value = "0">
			<input type="hidden" name='PositionI' id='Floor' value = "1">
			<input type="hidden" name='Side' id='Side' value = "Front">
			
			
		</div>
		<input type="submit" value = "Enter to Buggy (Sash Kit)" class="greenButton" onclick="Shift.action='ShifthardwareSashkit.asp'; DisableButton(this);"></input>
		<input type="submit" value = "Enter to Buggy (Frame Kit)" class="greenButton" onclick="Shift.action='ShifthardwareFramekit.asp'; DisableButton(this);"></input>
		<input type="submit" value = "View All" class="redButton" onclick="Shift.action='ShifthardwareViewAll.asp'; DisableButton(this);"></input>
</fieldset>
       <div>
        <h1>Sample Bin Cart View</h1>
		<p>
		<table>
		<tr><th>Front</th><th>Back</th></tr>
		<tr><td BGColor ="Yellow">
		<table border='1' class='sortable'>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		</table>
		</td>
		<td>
		<table border='1' class='sortable'>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		<tr><td BGColor ="Yellow">1</td><td BGColor ="Yellow">2</td><td BGColor ="Yellow">3</td><td BGColor ="Cyan">4</td><td BGColor ="Cyan">5</td><td BGColor ="Cyan">6</td><td BGColor ="Yellow">7</td><td BGColor ="Yellow">8</td><td BGColor ="Yellow">9</td><td BGColor ="Cyan">10</td><td BGColor ="Cyan">11</td><td BGColor ="Cyan">12</td></tr>
		<tr><td BGColor ="Cyan">1</td><td BGColor ="Cyan">2</td><td BGColor ="Cyan">3</td><td BGColor ="Yellow">4</td><td BGColor ="Yellow">5</td><td BGColor ="Yellow">6</td><td BGColor ="Cyan">7</td><td BGColor ="Cyan">8</td><td BGColor ="Cyan">9</td><td BGColor ="Yellow">10</td><td BGColor ="Yellow">11</td><td BGColor ="Yellow">12</td></tr>
		</table>
		</td></tr>
		</table>
		 </p>        
</div>      
              
    </ul>    
</form>	
             
</body>
</html>
