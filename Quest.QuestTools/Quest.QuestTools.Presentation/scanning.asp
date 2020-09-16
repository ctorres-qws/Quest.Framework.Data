<!--#include file="adovbs.inc"-->
<%

tablename=jobname
fu=floor
etag=tag

' This is the variable for offset '
' This could be collected from the previous page "

pagebreak="<P CLASS=""classPageBreak""><BR></P>"
dim pagecount1, pagecount2, pagecount3

' The include for the connection file below '


%>
<!--#include file="connect_scanning.asp"-->
<html><head><title>:: Quest Window Systems Inc. :: Job Name =<%=tablename%> Floor =<%=fu%></title></head>
<body topmargin="0" leftmargin="0" link="#000000" vlink="#C0C0C0" alink="#F6F000">

<%
rs.filter = "Tag = " & etag

showlabel = 0 ' show the label
strTag = ""
strHeight = ""

%>

<%
DIM intwidth, intwidth2, mulwidth, mulwidth2, IH1, IH2, IH3, IH4, IH5, IH6, IH7, IH8
DIM sty



rs.close
set rs=nothing
rs2.close
set rs2=nothing
DBConnection.close
set DBConnection=nothing
%>
</body></html>