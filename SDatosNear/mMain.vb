Module mMain

    Public Sub Main()
        Call OpenConnections()
    End Sub

    Public Sub GlobalErrorHandler(ByVal ex As Exception, Optional ByVal pvstr_Mensaje As String = "")
        If System.Environment.CommandLine.Split("/").Length > 0 OrElse Trim(System.Environment.CommandLine.Split("/")(1)) <> "AUTORUN" Then
            Dim loFrm As New frmFatalError
            loFrm.ShowCritical(ex, pvstr_Mensaje)
            loFrm = Nothing
        End If
    End Sub

End Module
