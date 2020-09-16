<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<%
'Put this is a code behind module or asp.net page
Sub DisplayDownloadDialog(ByVal PathVirtual As String)

        Dim strPhysicalPath As String
        Dim objFileInfo As System.IO.FileInfo
        Try
            strPhysicalPath = Server.MapPath(PathVirtual)
            'exit if file does not exist
            If Not System.IO.File.Exists(strPhysicalPath) _
                 Then Exit Sub
            objFileInfo = New System.IO.FileInfo(strPhysicalPath)
            Response.Clear()
           'Add Headers to enable dialog display
            Response.AddHeader("Content-Disposition", "attachment; filename=" & _
                objFileInfo.Name)
            Response.AddHeader("Content-Length", objFileInfo.Length.ToString())

            Response.ContentType = "application/octet-stream"
            Response.WriteFile(objFileInfo.FullName)


        Catch
            'on exception take no action
            'you can implement differently
        Finally

            Response.End()

        End Try
    End Sub
'DEMO
'IN YOUR CODE BEHIND PAGE:
'DECLARATION
Protected WithEvents btnDownload As _
  System.Web.UI.WebControls.Button

'IN CODE
'ASSUMES MYWORDFILE.DOC EXISTS IN SAME FOLDER
'AS THE ASPX FILE
 Private Sub btnDownload_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) _
          Handles btnDownload.Click
        DisplayDownloadDialog("MyWordFile.doc")
 End Sub
'ON APSX PAGE

%>
<body>
</body>
</html>
