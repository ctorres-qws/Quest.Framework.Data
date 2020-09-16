<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Password Central Location for Export-->
<!-- Export Buttons are 1 hit only, if hit accidently, database must be corrected -->
<!-- This Password ensures that only clicking with a password can activate it  -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optima Password</title>
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

<%
Passkey = "SASHA"
Password = UCASE(TRIM(Request.Form("pwd")))

ExportSite = REQUEST.QueryString("ExportSite")

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a <a class="button leftButton" type="cancel" href="Index.html#_Glass" target="_self">Glass Tool</a>

    </div>
<%
If Password = Passkey Then

	Select Case ExportSite
		Case "Production"
			ExportSiteAddress = "GlassExportProduction.asp"
			Response.Redirect ExportSiteAddress
		Case "Service"
			ExportSiteAddress = "GlassExportService.asp"
			Response.Redirect ExportSiteAddress
		Case "Commercial"
			ExportSiteAddress = "GlassExportCommercial.asp"
			Response.Redirect ExportSiteAddress
		Case "All"
			ExportSiteAddress = "GlassExport.asp"
			Response.Redirect ExportSiteAddress
		Case "Safe"
		' Password removed from Safe Case but can still be used as an example
			ExportSiteAddress = "GlassOptimaSafe.asp"
			Response.Redirect ExportSiteAddress
		Case Else
			ExportSiteAddress = "Index.Html#_Glass"
			Response.Redirect ExportSiteAddress
	End Select

Else
%>
<form id="adminpass" title="Optima Export Password" class="panel" name="enter" action="GlassExportPassword.asp?ExportSite=<%response.write ExportSite%>" method="post" target="_self" selected="True">
<H2><% response.write ExportSite %> </H2>
<fieldset>
			<div class="row" >
				
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>

<%
End If
%>

<%
If Password = Passkey Then
%>

<%
End If
%>

</body>
</html>

<%
DBConnection.close
set DBConnection=nothing
%>

