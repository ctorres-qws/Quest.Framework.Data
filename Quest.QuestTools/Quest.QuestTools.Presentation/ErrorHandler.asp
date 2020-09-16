<%

	Dim str_Err
	
	str_Err = str_Err & Server.GetLastError()

	
	Response.Write(str_Err)

%>