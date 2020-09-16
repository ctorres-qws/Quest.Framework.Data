<!--#include file="adovbs.inc"-->
<!--#include file="inc_style.asp"-->
<%
' This is the variable for offset '
' This could be collected from the previous page "

%>
<%
job_name=request.querystring("jobname")
floor_number=request.querystring("fl")
pono=request.querystring("po")

' The include for the connection file below '

%>
<!--#include file="connect_supp2.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<title>Home Page</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body link="#000000" vlink="#C0C0C0" alink="#F6F000">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="67%" id="AutoNumber2">
  <tr>
    <td width="57%"><font face="Arial"><b>Quest Windows Systems, Inc.</b></font></td>
    <td width="43%">
    <p align="right"><b><font face="Arial" size="5"><br>
    </font><font face="Arial" size="2">Order Report</font></b></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" bgcolor="#C0C0C0" bordercolor="#FFFFFF">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="652" id="AutoNumber3">
      <tr>
        <td width="86"><b>
    <font face="Arial" size="4">JOB: <% Response.Write rs("Job") %> </font>
    
 </b>
    
        </td>
        <td width="112"><b>
    <font face="Arial" size="4">FLOOR: <% Response.Write rs("Floor") %></font></b></td>
        <td width="197"><font face="Arial" size="4"><b>Master P.O. # <% Response.Write rs2("MasterPO") %></b></font></td>
        <td width="257"><b><font face="Arial" size="4">Order</font></b><font face="Arial" size="4"><b> # <% Response.Write rs2("PO") %></b></font></td>
      </tr>
      </table>
    
    </td>
  </tr>
</table>
<img border="0" src="window/spacer.gif" WIDTH="1" HEIGHT="1">

<hr>
<% 
rs.filter = "Style <= 0"
 %>

<% Do While Not rs.eof %>
<font face='Arial' size='2'>
<table border="1" style="border-collapse: collapse" bordercolor="#C0C0C0" backgroundcolor="FFFFFF" width="100%" id="AutoNumber1" cellpadding="2">
<tr>
<td width='11' bgcolor="#3366FF" valign="top" bordercolor="#C0C0C0"><b>
<font size="1" color="#FFFFFF">SU</font></b></td>
<td width='81' bgcolor="#F2F2F2" valign="top"><font size="2">QTY</font><font size="2" face="Arial"><B>:
<% Response.Write rs("Qty") %></B></font></td>
<td width='123' bgcolor="#F2F2F2" valign="top"><font size="2">W : </font>
<font face="Arial" size="2"><b><% Response.Write rs("W")%>"</td></b></font>
<td width='123' bgcolor="#F2F2F2" valign="top"><font size="2">H : </font>
<font face="Arial" size="2"><b><% Response.Write rs("H")%>"</td></b></font></font>
<td width='755' bgcolor="#F2F2F2" valign="top"><font face="Arial" size="2">Tags:<b><% Response.Write rs("Tag")%></font></td></b></font>
  </b></font>
<td width='109' bgcolor="#F2F2F2" valign="top">
<font face="Arial Narrow" size="2">Overall: <% Response.Write rs("Overall")%>&quot;</font></td>
</tr></table>
<img border="0" src="window/spacer.gif" WIDTH="1" HEIGHT="1">
        </b></b>
<% rs.MoveNext %>
<% loop
rs.close
 %><HR>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="67%" id="AutoNumber4">
  <tr>
    <td width="100%" bgcolor="#F2F2F2"><b><font face="Arial" color="#FF6600">
    TEMPERED GLASS</font></b><font face="Arial"><b>: <% response.write rs2("TGNotes") %></b></font></td>
  </tr>
</table><hr>
<% 
rs.open
rs.filter = "Style > 0" %>
<% Do While Not rs.eof %>
<font face='Arial' size='2'>
<table border="1" style="border-collapse: collapse" bordercolor="#C0C0C0" backgroundcolor="FFFFFF" width="100%" id="AutoNumber1" cellpadding="2">
<tr>
<td width='5' bgcolor="#FF6600" valign="top"><b><font size="1" color="#FFFFFF">TG</font></b></td>
<td width='84' bgcolor="#F2F2F2" valign="top"><font size="2">QTY</font><font size="2" face="Arial"><B>:
<% Response.Write rs("Qty") %></B></font></td>
<td width='119' bgcolor="#F2F2F2" valign="top"><font size="2">W : </font>
<font face="Arial" size="2"><b><% Response.Write rs("W")%>"</td></b></font>
<td width='121' bgcolor="#F2F2F2" valign="top"><font size="2">H : </font>
<font face="Arial" size="2"><b><% Response.Write rs("H")%>"</td></b></font></font>
<td width='757' bgcolor="#F2F2F2" valign="top"><font face="Arial" size="2">Tags:<b><% Response.Write rs("Tag")%></font></td></b></font>
  </b></font>
<td width='111' bgcolor="#F2F2F2" valign="top">
<font face="Arial Narrow" size="2">Overall: <% Response.Write rs("Overall")%>&quot;</font></td>
</tr></table>
<img border="0" src="window/spacer.gif" WIDTH="1" HEIGHT="1">
        </b></b>
<% rs.MoveNext %><% loop %><%
rs.close
set rs=nothing
rs2.close
set rs2=nothing
DBConnection.close
set DBConnection=nothing
%></table></font>