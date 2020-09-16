<%
tablename=jobname
fl1=floor
etag=tag
%>

<!--#include file="connect_flashbc.asp"-->

<%
rs.filter = "Tag = " & etag


if rs("Y") > 110 then
factor = 3
else
factor = 4
end if

ih1 = rs("H1") * factor
ih2 = (rs("H2") * factor) + ih1
ih3 = (rs("H3") * factor) + ih2
ih4 = (rs("H4") * factor) + ih3
ih5 = (rs("H5") * factor) + ih4
ih6 = (rs("H6") * factor) + ih5
ih7 = (rs("H7") * factor) + ih6
ih8 = (rs("H8") * factor) + ih7
lm0 = rs("LM0") * factor
lm1 = rs("LM1") * factor
rm0 = rs("RM0") * factor
rm1 = rs("RM1") * factor
width = rs("X") * factor
hwidth = width - 2
height = ( rs("Y") -1 ) * factor
y2 = rs("Y2") * factor
vw = rs("VW") * factor
vh = rs("VH") * factor

'vars just for flash'

hhwidth = hwidth / 2
hhhwidth = hhwidth / 2
vlr = hhwidth + vw
vll = hwidth - vw
vlll = hhwidth - vw
hhhhwidth = hwidth - hhhwidth

hh1 = (rs("H1") * factor) / 2
hh2 = (rs("H2") * factor) / 2 + ih1
hh3 = (rs("H3") * factor) / 2 + ih2
hh4 = (rs("H4") * factor) / 2 + ih3 
hh5 = (rs("H5") * factor) / 2 + ih4 
hh6 = (rs("H6") * factor) / 2 + ih5
hh7 = (rs("H7") * factor) / 2 + ih6
hh8 = (rs("H8") * factor) / 2 + ih7

hlm0 = (rs("LM0") * factor) /2
hlm1 = (rs("LM1") * factor) /2 + lm0
hrm0 = (rs("RM0") * factor) /2
hrm1 = (rs("RM1") * factor) /2 + rm0



truex = rs("X")
truey = rs("Y")


style = rs("Style")

rs2.filter = "Name = " & style
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

%><%
whichFN=server.mappath("/testflash.txt")


' first, create the file out of thin air
Set fstemp = server.CreateObject("Scripting.FileSystemObject")
Set filetemp = fstemp.CreateTextFile(whichFN, true)
' true = file can be over-written if it exists
' false = file CANNOT be over-written if it exists

filetemp.WriteLine("y2=" & y2 & "&height=" & height & "&width=" & width & "&hwidth=" & hwidth & "&hhwidth=" & hhwidth & "&hhhwidth=" & hhhwidth & "&hhhhwidth=" & hhhhwidth & "&ih1=" & ih1 & "&ih2=" & ih2 & "&ih3=" & ih3 & "&ih4=" & ih4 & "&ih5=" & ih5 & "&ih6=" & ih6 & "&ih7=" & ih7 & "&ih8=" & ih8 & "&lm0=" & lm0 & "&lm1=" & lm1 & "&rm0=" & rm0 & "&rm1=" & rm1 & "&hlm0=" & hlm0 & "&hlm1=" & hlm1 & "&hrm0=" & hrm0 & "&hrm1=" & hrm1 & "&truex=" & truex & "&truey=" & truey & "&o1=" & o1 & "&o2=" & o2 & "&o3=" & o3 & "&o4=" & o4 & "&o5=" & o5 & "&o6=" & o6 & "&o7=" & o7 & "&o8=" & o8 & "&r0=" & r0 & "&r1=" & r1 & "&r2=" & r2 & "&r3=" & r3 & "&r4=" & r4 & "&r5=" & r5 & "&r6=" & r6 & "&r7=" & r7 & "&r8=" & r8 & "&l0=" & l0 & "&l1=" & l1 & "&l2=" & l2 & "&l3=" & l3 & "&l4=" & l4 & "&l5=" & l5 & "&l6=" & l6 & "&l7=" & l7 & "&l8=" & l8 & "&vw=" & vw & "&vh=" & vh & "&vlr=" & vlr & "&vll=" & vll & "&vlll=" & vlll & "&hh1=" & hh1 & "&hh2=" & hh2 & "&hh3=" & hh3 & "&hh4=" & hh4 & "&hh5=" & hh5 & "&hh6=" & hh6 & "&hh7=" & hh7 & "&hh8=" & hh8 & "&bullshit=0" )
strTemp = "y2=" & y2 & "&height=" & height & "&width=" & width & "&hwidth=" & hwidth & "&hhwidth=" & hhwidth & "&hhhwidth=" & hhhwidth & "&hhhhwidth=" & hhhhwidth & "&ih1=" & ih1 & "&ih2=" & ih2 & "&ih3=" & ih3 & "&ih4=" & ih4 & "&ih5=" & ih5 & "&ih6=" & ih6 & "&ih7=" & ih7 & "&ih8=" & ih8 & "&lm0=" & lm0 & "&lm1=" & lm1 & "&rm0=" & rm0 & "&rm1=" & rm1 & "&hlm0=" & hlm0 & "&hlm1=" & hlm1 & "&hrm0=" & hrm0 & "&hrm1=" & hrm1 & "&truex=" & truex & "&truey=" & truey & "&o1=" & o1 & "&o2=" & o2 & "&o3=" & o3 & "&o4=" & o4 & "&o5=" & o5 & "&o6=" & o6 & "&o7=" & o7 & "&o8=" & o8 & "&r0=" & r0 & "&r1=" & r1 & "&r2=" & r2 & "&r3=" & r3 & "&r4=" & r4 & "&r5=" & r5 & "&r6=" & r6 & "&r7=" & r7 & "&r8=" & r8 & "&l0=" & l0 & "&l1=" & l1 & "&l2=" & l2 & "&l3=" & l3 & "&l4=" & l4 & "&l5=" & l5 & "&l6=" & l6 & "&l7=" & l7 & "&l8=" & l8 & "&vw=" & vw & "&vh=" & vh & "&vlr=" & vlr & "&vll=" & vll & "&vlll=" & vlll & "&hh1=" & hh1 & "&hh2=" & hh2 & "&hh3=" & hh3 & "&hh4=" & hh4 & "&hh5=" & hh5 & "&hh6=" & hh6 & "&hh7=" & hh7 & "&hh8=" & hh8 & "&bullshit=0"
filetemp.writeblanklines(3) 
filetemp.Close

%>
<body topmargin="0" leftmargin="0">
<!--cahnges for refresh cache #sachin #date 10 feb-->

<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width="240" height="480" id="styles_preview" align="middle">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="styles_preview.swf?<%=strTemp%>" />
<param name="quality" value="high" />
<param name="bgcolor" value="#ffffff" />
<embed src="styles_preview.swf?<%=strTemp%>" quality="high" bgcolor="#ffffff" width="240" height="480" name="styles_preview" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object>

</body>