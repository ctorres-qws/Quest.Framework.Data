<!-- #include virtual="/readini/inifile.inc" -->
<html>

<head>
<title>INI File Reading Test</title>
</head>

<body>

<p>Read the INI file (Physical path)...<%
  call IniFileLoad("physical=c:\boot.ini")
%></p>

<p>Dump the dictionary:<%
  dim TempArray
  TempArray = IniFileDictionary.Keys
  for i = 0 to IniFileDictionary.Count - 1
    Response.Write("<br>" & TempArray(i) & "=" & IniFileDictionary(TempArray(i)))
  next
%></p>

<p>Display certain values:<%
  StrBuf = IniFileValue("boot loader|timeout")
  Response.Write("<br>'[boot loader] timeout' value = " & StrBuf)
  StrBuf = IniFileValue("boot loader")
  Response.Write("<br>'[boot loader]' section = " & StrBuf)
%></p>

<p>Read the INI file (Virtual path)...<%
  call IniFileLoad("virtual=/PL_WIDTH/DC500.ini")
%></p>

<p>Dump the dictionary:<%
  TempArray = IniFileDictionary.Keys
  for i = 0 to IniFileDictionary.Count - 1
    Response.Write("<br>" & TempArray(i) & "=" & IniFileDictionary(TempArray(i)))
  next
%></p>

<p>Display certain values:<%
  StrBuf = IniFileValue("names|name2")
  Response.Write("<br>'[names] name2' value = " & StrBuf)
  StrBuf = IniFileValue("colors")
  Response.Write("<br>'[colors]' section = " & StrBuf)
%></p>
</body>
</html>
