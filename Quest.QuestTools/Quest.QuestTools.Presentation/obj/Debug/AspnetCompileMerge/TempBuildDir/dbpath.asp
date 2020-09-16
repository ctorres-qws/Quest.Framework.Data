<!-- DBPath creates Database Connection to Quest Database and recognizes ACCESS AND SQL-->
<!-- February 2019 - New Addition to this page USA/CANADA Marker "CountryLocation"-->
<!--#include file="@common.asp"-->

<%

'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp

'Added Michael Bernholtz November 2018
DBConnection.ConnectionTimeout = 1000
DBConnection.CommandTimeout = 1000
'End add

DBConnection.Open DSN

' Location Marker Include
%>
<!--#include file="CountryLocation.INC"-->



	
