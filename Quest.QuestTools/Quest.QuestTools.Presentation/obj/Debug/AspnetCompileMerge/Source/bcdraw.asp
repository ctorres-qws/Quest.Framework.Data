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
<!--#include file="connect_flashbc.asp"-->
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

tag = rs("Tag")
style = rs("Style")
Awidth = rs("X")
Aheight = rs("Y")
y2 = rs("Y2")

H1 = rs("H1")
H2 = rs("H2")
H3 = rs("H3")
H4 = rs("H4")
H5 = rs("H5")
H6 = rs("H6")
H7 = rs("H7")
H8 = rs("H8")
lm0= rs("LM0")
lm1= rs("LM1")
rm0= rs("RM0")
rm1= rs("RM1")
vw = rs("VW")
vh = rs("VH")

'intwidth = rs("X") - 2'
'mulwidth = rs("X") - 3'
'mulwidth = mulwidth / 2'

sty = rs("Style")
rs2.filter = "Name = " & sty

o1 = rs2("O1")
o2 = rs2("O2")
o3 = rs2("O3")
o4 = rs2("O4")
o5 = rs2("O5")
o6 = rs2("O6")
o7 = rs2("O7")
o8 = rs2("O8")
l0 = rs2("L0")
l1 = rs2("L1")
l2 = rs2("L2")
l3 = rs2("L3")
l4 = rs2("L4")
l5 = rs2("L5")
l6 = rs2("L6")
l7 = rs2("L7")
l8 = rs2("L8")
r0 = rs2("R0")
r1 = rs2("R1")
r2 = rs2("R2")
r3 = rs2("R3")
r4 = rs2("R4")
r5 = rs2("R5")
r6 = rs2("R6")
r7 = rs2("R7")
r8 = rs2("R8")

movieht = Aheight*2.5+80

if (strcomp(trim(strTag),left(tag,len(tag)-1))=0) and (strcomp(trim(strHeight),H1 & "," & H2 & "," & H3 & "," & H4 & "," & H5 & "," & H6 & "," & H7 & "," & H8)=0) then
	showlabel = 1 ' do not show the label
	moviewd = Awidth*2.5+10
	if moviewd < 101 then moviewd = 101
		'Response.Write "YYY"
else
	showlabel = 0 'show the label
	moviewd = Awidth*2.5+100
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
end if
%>

<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="<%=moviewd%>" height="<%=movieht%>" align="middle">
  <param name="movie" value="print_page.swf?showlabel=<%=showlabel%>&tag=<%=tag%>&sty=<%=sty%>&Awidth=<%=Awidth%>&Aheight=<%=Aheight%>&y2=<%=y2%>&H1=<%=H1%>&H2=<%=H2%>&H3=<%=H3%>&H4=<%=H4%>&H5=<%=H5%>&H6=<%=H6%>&H7=<%=H7%>&H8=<%=H8%>&lm0=<%=lm0%>&lm1=<%=lm1%>&rm0=<%=rm0%>&rm1=<%=rm1%>&vw=<%=vw%>&vh=<%=vh%>&o1=<%=o1%>&o2=<%=o2%>&o3=<%=o3%>&o4=<%=o4%>&o5=<%=o5%>&o6=<%=o6%>&o7=<%=o7%>&o8=<%=o8%>&l0=<%=l0%>&l1=<%=l1%>&l2=<%=l2%>&l3=<%=l3%>&l4=<%=l4%>&l5=<%=l5%>&l6=<%=l6%>&l7=<%=l7%>&l8=<%=l8%>&r0=<%=r0%>&r1=<%=r1%>&r2=<%=r2%>&r3=<%=r3%>&r4=<%=r4%>&r5=<%=r5%>&r6=<%=r6%>&r7=<%=r7%>&r8=<%=r8%>">
  <param name="quality" value="high">
  <param name="menu" value="false">
  <embed src="print_page.swf?showlabel=<%=showlabel%>&tag=<%=tag%>&sty=<%=sty%>&Awidth=<%=Awidth%>&Aheight=<%=Aheight%>&y2=<%=y2%>&H1=<%=H1%>&H2=<%=H2%>&H3=<%=H3%>&H4=<%=H4%>&H5=<%=H5%>&H6=<%=H6%>&H7=<%=H7%>&H8=<%=H8%>&lm0=<%=lm0%>&lm1=<%=lm1%>&rm0=<%=rm0%>&rm1=<%=rm1%>&vw=<%=vw%>&vh=<%=vh%>&o1=<%=o1%>&o2=<%=o2%>&o3=<%=o3%>&o4=<%=o4%>&o5=<%=o5%>&o6=<%=o6%>&o7=<%=o7%>&o8=<%=o8%>&l0=<%=l0%>&l1=<%=l1%>&l2=<%=l2%>&l3=<%=l3%>&l4=<%=l4%>&l5=<%=l5%>&l6=<%=l6%>&l7=<%=l7%>&l8=<%=l8%>&r0=<%=r0%>&r1=<%=r1%>&r2=<%=r2%>&r3=<%=r3%>&r4=<%=r4%>&r5=<%=r5%>&r6=<%=r6%>&r7=<%=r7%>&r8=<%=r8%>" width="<%=moviewd%>" height="<%=movieht%>" align="middle" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"></embed></object>

<%

strTag = left(tag,len(tag)-1)
'Response.Write "<BR>" & strTag & "<BR>"
strHeight = H1 & "," & H2 & "," & H3 & "," & H4 & "," & H5 & "," & H6 & "," & H7 & "," & H8
'Response.Write strHeight & "<BR>"


rs.close
set rs=nothing
rs2.close
set rs2=nothing
DBConnection.close
set DBConnection=nothing
%>
</body></html>