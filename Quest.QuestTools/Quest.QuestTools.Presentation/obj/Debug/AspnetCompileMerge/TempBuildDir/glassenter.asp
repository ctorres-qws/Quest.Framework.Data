<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Edited by Michael Bernholtz during December 2013 to add new fields-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Glass</title>
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
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

              <form id="enter" title="Enter New Glass Form" class="panel" name="enter" action="glassconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter New Glass Information:</h2>

                                <fieldset>
  <div class="row">

                        <div class="row">
                <label>Job </label>
                <input type="text" name='PROJECT' id='PROJECT' >
            </div>

    <div class="row">
                <label>Floor</label>
                <input type="text" name='FLOOR' id='FLOOR' >
            </div>

        <div class="row">
                <label>Tag</label>
                <input type="text" name='TAG' id='TAG' >
        </div>
<!-- Added on December 5th - Department choices hardcoded -->
		<div class="row">
                <label>Department</label>
                <select name= 'DEPARTMENT' id = 'DEPARTMENT'>
					<option value="Production">Production</option>
					<option value="Service">Service</option>
					<option value="Commercial">Commercial</option>
					<option value="Recut">Recut</option>
				</select>
        </div>

            <div class="row">
                <label>Notes</label>
                <input type="text" name='NOTES' id='NOTES' >
            </div>

            <div class="row">
                <label>Ext.Glass</label>
                <select name="ONEMAT">
<% mat = mat1 %>
<!--#include file="QSU.inc"-->
</select>
                 </div>
		<div class="row">
                <label>EXT Method</label>
				<select name ='EXTMethod'>
					<option value = 'CUT' selected >CUT</option>
					<option value = 'ORDER'>ORDER</option>
					<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
                </select>
            </div>
      
            
              <div class="row">
                <label>Spacer</label>
                <select name="ONESPAC">
<% mat = spac1 %>
<!--#include file="QSU2.inc"-->
</select>
            </div>

			 <div class="row">
                <label>Black/Grey</label>
				<select name ='SPACERCOLOUR'>
					<option value = 'Black' selected >Black</option>
					<option value = 'Grey'>Grey</option>
                </select>
            </div>

            <div class="row">
                <label>Int.Glass</label>
                <select name="TWOMAT">
<% mat = mat1 %>
<!--#include file="QSU.inc"-->
</select>
                 </div>
				 		<div class="row">
                <label>INT Method</label>
				<select name ='INTMethod'>
					<option value = 'CUT' selected >CUT</option>
					<option value = 'ORDER'>ORDER</option>
					<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
                </select>
            </div>

              <div class="row">
                <label>AIR / ARGON</label>
				<select name ='AIR'>
					<option value = 'Argon' selected >Argon</option>
					<option value = 'Air'>Air</option>
					<option value = 'N/A'>N/A</option>
                </select>
            </div>

              <div class="row">
                <label>Width</label>
                <input type="text" name='WIDTH' id='WIDTH' size='8'>
              </div>

              <div class="row">
                <label>Height</label>
                <input type="text" name='HEIGHT' id='HEIGHT' size='8'>
            </div>

			<div class="row">
                <label>Order By</label>
				<select name ='orderBy'>
				<option value = 'Yegor'>Yegor</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'Joe'>Joe</option>
				<option value = 'Ariel'>Ariel</option>
				<option value = 'Hamid'>Hamid</option>
				<option value = 'Michael'>Michael</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Sasha'>Sasha</option>
				<option value = 'John'>John</option>
				<option value = 'WIS'>WIS</option>
                </select>

             </div>

			 <div class="row">
                <label>Order For</label>
				<select name ='orderFor'>
				<option value = 'Arten'>Artem</option>
				<option value = 'Daniel'>Daniel</option>
				<option value = 'Ellerton'>Ellerton</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'George'>George</option>
				<option value = 'Hamlet'>Hamlet</option>
				<option value = 'Ivan'>Ivan</option>
				<option value = 'John'>John</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Rob'>Rob</option>
				<option value = 'Roman'>Roman</option>
				<option value = 'Vince'>Vince</option>
				<option value = 'Yegor'>Yegor</option>
				<option value = 'WIS'>WIS</option>
                </select>

             </div>

			 <div class="row">
                <label>PO #</label>
                <input type="text" name='PoNum' id='PoNum'>
             </div>

			<div class="row">
                <label>Ext Work #</label>
                <input type="text" name='ExtorderNum' id='ExtorderNum' >
             </div>

			 <div class="row">
                <label>Ext Expected Date</label>
                <input type="date" name='ExtExpected' id='ExtExpected' >
             </div>

			 <div class="row">
                <label>Ext From</label>
                <select name ='ExtFrom'>
				<option value = 'Quest' selected>Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select>
             </div>

			<div class="row">
                <label>Int Work #</label>
                <input type="text" name='IntorderNum' id='IntorderNum' >
             </div>

			 <div class="row">
                <label>Int Expected Date</label>
                <input type="date" name='IntExpected' id='IntExpected' >
             </div>

			 <div class="row">
                <label>Int From</label>
				<select name ='IntFrom'>
				<option value = 'Quest' selected>Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select>
             </div>

<!-- Added December 4th - Require Date is currently a plain text field-->

              <div class="row">
                <label>Req. Date</label>
				<% 
				tenDay = DateAdd("d",10,Date()) 
				%>
                <input type="text" name='REQUIREDDATE' id='REQUIREDDATE' size='8' value='<% response.write tenDay %>' >
				<ul>
				<li>In the Format: MM/DD/YYYY </li>
				</ul>

            </div>
			<div class="row">
				<label>Glass For</label>
				<select name ='GlassFor'>
					<option value = 'SU'>SU - Sealed Unit</option>
					<option value = 'OV'>OV - Operable Vent</option>
					<option value = 'SP'>SP - Spandrel</option>
					<option value = 'SB'>SB -Shadow Box</option>
					<option value = 'SBOC'>SBOC - Shadow Box Outside Corner</option>
					<option value = 'SBIC'>SBIC - Shadow Box Inside Corner</option>
					<option value = 'SW'>SW - Swing Door/option>
					<option value = 'SD'>SD - Sliding Door</option>
					<option value = 'Sunview Door'>Sunview Door</option>
					<option value = 'OC'>OC - Outside Corner Offset</option>
					<option value = 'IC'>IC - Inside Corner Offset</option>
					<option value = 'DOS'>DOS -Double Offset</option>
				</select>
			</div>
			

				<div class="row">
				<label>SP Colour</label>
				<select name ='SPColour'>
				<option value = ''></option>
				<option value = 'SP1'>SP1</option>
				<option value = 'SP2'>SP2</option>
				<option value = 'SP3'>SP3</option>
				<option value = 'SP4'>SP4</option>
				<option value = 'SP5'>SP5</option>
				<option value = 'SP6'>SP6</option>
				<option value = 'SP7'>SP7</option>
				<option value = 'SP8'>SP8</option>
                </select>
				</div>

                    <a class="whiteButton" href="javascript:enter.submit()">Submit</a>

</fieldset>

            </form>

<%
DBConnection.close
Set DBConnection = nothing
%>

</body>
</html>
