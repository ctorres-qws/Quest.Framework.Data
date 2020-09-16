<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--Optima Selection Page for adding QT FILE, Sasha just wants to enter ids-->
		<!--Created January 2015, at Request of Sasha for adding a note to multiple items at once-->
		<!-- Sends to glassOptimaQTConf.asp-->

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
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       <form id="Optima" action="glassOptimaNoteConf2.asp" name="Optima"  method="GET" target="_self" selected="true" >  
        
		<h2><center>Add Key Details to many records all at once by inputting each ID separated by a comma<center></h2>
		<h2><center>For Quest Glass please put MO number into the Ext/Int Work # Field<center></h2>
		<h2><center>For Externally Ordered Glass please put Order Number into the Ext/Int Work # Field<center></h2>
		
		<fieldset>


		<div class="row">
			<label>Add Note</label>
			<input type="text" name='NOTES' id='NOTES' />
		</div>
		<div class="row">
			<label>Window PO</label>
			<input type="text" name='PoNum' id='PoNum' >
        </div>
			 
		<div class="row">
			<label>Ext Work #</label>
			<input type="text" name='ExtorderNum' id='ExtorderNum' >
        </div>
		<div class="row">
			<label>Int Work #</label>
			<input type="text" name='IntorderNum' id='IntorderNum' >
		</div>		
		<div class="row">
			<label>QT File</label>
			<input type="text" name='QTFile' id='QTFile' >
        </div>
		<div class="row">
                <label>Glass IDs</label>
                <input type="text" name='IDList' id='IDList' />
        </div>
               <input type="hidden" name='ticket' id='ticket' value = 'multiple' />     
		<a class="whiteButton" onClick="Optima.action='GlassOptimaNoteConf2.asp'; Optima.submit()">ADD QT</a><BR>
		</fieldset>
        <ul id="Profiles" title=" Optima Report" selected="true">
	

	</table>
	
      </ul>    
		</form>
            
            
            
       
            
              
               
                
             
               
</body>
</html>
