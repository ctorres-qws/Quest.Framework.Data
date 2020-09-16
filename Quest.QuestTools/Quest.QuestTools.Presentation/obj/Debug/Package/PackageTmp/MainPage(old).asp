
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
	<title>Quest Tools Redirect</title>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no" />

</head>
<body>	 
<%

IPADDRESS = Request.ServerVariables("REMOTE_ADDR")
PORT = Request.ServerVariables("SERVER_PORT")

TexasPort = "8082"
CanadaPort = "8081"

TexasLowIPRange = "10.34.16.1"
TexasHighIPRange = "10.34.31.254"

'TexasLowIPRange = "192.168.0.0"
'TexasHighIPRange = "192.168.254.254"


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

If (ip2num(IPADDRESS) >= ip2num(TexasLowIPRange) And ip2num(IPADDRESS) <= ip2num(TexasHighIPRange) )or ip2num(IPADDRESS) = ip2num("192.168.2.143") or ip2num(IPADDRESS) = ip2num("192.168.2.58") Then
	CountryLocation = "USA"
Else
	CountryLocation = "CANADA" 
End If	


		
		
IF CountryLocation = "USA" then

	%>
	<meta http-equiv="refresh" content="0;url='indexTexas.html" />
	<%
Else
	%>
	<meta http-equiv="refresh" content="0;url='index.html" />
	<%
End if
%>



</body>
</html>