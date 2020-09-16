<%
'Include for IPADDRESS - February 2019
'Michael Bernholtz

IPADDRESS = Request.ServerVariables("REMOTE_ADDR")
PORT = Request.ServerVariables("SERVER_PORT")

TexasPort = "8082"
CanadaPort = "8081"

' TexasLowIPRange = "10.34.16.1"  Low Range Given by ETS
' TexasHighIPRange = "10.34.31.254" High Range Given by ETS
'TexasLowIPRange = "10.34.0.0"
'TexasHighIPRange = "10.34.255.255"
TexasLowIPRange = "10.34.22.0"
TexasHighIPRange = "10.34.31.255"

Texas1 = ip2num("192.168.0.219") 'Bernholtz wired
Texas2 = ip2num("192.168.0.154") 'Bernholtz wifi
Texas3 = ip2num("10.34.26.17") 'Morgan wirelss
Texas4 = ip2num("10.34.25.34") 'Morgan wifi
Texas5 = ip2num("10.34.26.17") 'placeholder
Canada0 = ip2num("10.34.1.42") '172.18.13.31
Canada1 = ip2num("10.34.26.16") 'Kevin in texas
Canada2 = ip2num("10.34.3.199") ' Jennifer VPN
Canada3 = ip2num("10.34.3.29") 'Iden VPN
Canada4 = ip2num("10.34.3.217") 'Diane Cook VPN
Canada5 = ip2num("10.34.3.84") 'Iden VPN 2
Canada6 = ip2num("10.34.24.2") 'Liam
Canada7 = ip2num("10.34.3.114") 'Liam
Canada8 = ip2num("10.34.3.176") 'Liam
Canada9 = ip2num("10.34.26.13") 'Clint work from Texas
Canada10 = ip2num("10.34.3.89") ' Jennifer VPN

CountryLocation = 	 ""
'CountryLocation by port
'If PORT = TexasPort then
'	CountryLocation = "USA"
'else
'	CountryLocation = "CANADA"
'end if 

Public Function ip2num(ip)
Dim i, a, N
a = Split(ip, ".")
N = CDbl(0)
For i = 0 To UBound(a)
  N = N * 256 + a(i)
Next 
ip2num = N
End Function

'  
'Michael Bernholtz and Ariel Aziza Quest Tools included for Testing
If ( (ip2num(IPADDRESS) >= ip2num(TexasLowIPRange) And ip2num(IPADDRESS) <= ip2num(TexasHighIPRange)) and ip2num(IPADDRESS) <> Canada0 and ip2num(IPADDRESS) <> Canada1 and  ip2num(IPADDRESS) <> Canada2 and  ip2num(IPADDRESS) <> Canada3 and  ip2num(IPADDRESS) <> Canada4 and  ip2num(IPADDRESS) <> Canada5 and  ip2num(IPADDRESS) <> Canada6 and  ip2num(IPADDRESS) <> Canada7 and ip2num(IPADDRESS) <> Canada8 and  ip2num(IPADDRESS) <> Canada9 and  ip2num(IPADDRESS) <> Canada10 ) Then

	CountryLocation = "USA"

Else
	CountryLocation = "CANADA" 
End If	

'Michael Bernholtz and Ariel Aziza Quest Tools included for Testing
If (ip2num(IPADDRESS) = Texas1 or ip2num(IPADDRESS) = Texas2 or ip2num(IPADDRESS) = Texas3 ) or ip2num(IPADDRESS) = Texas4 or ip2num(IPADDRESS) = Texas5 THEN
	CountryLocation = "USA"
End If	
%>