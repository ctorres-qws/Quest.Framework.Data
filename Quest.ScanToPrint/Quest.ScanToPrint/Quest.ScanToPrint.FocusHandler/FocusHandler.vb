Imports System.Windows.Forms

Public Module FocusHandler
    Public Sub SetFocus(ByVal serviceName)
        If Process.GetProcessesByName(serviceName).Length >= 1 Then
            For Each ObjProcess As Process In Process.GetProcessesByName(serviceName)
                AppActivate(ObjProcess.Id)

                SendKeys.SendWait("~")
            Next
        End If
    End Sub
End Module
