Public Class cMail
    Public Shared Sub SendMail( _
        pvstrTo As String, _
        pvstrFrom As String, _
        pvstrSubject As String, _
        pvstrBody As String)

        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient

        'CONFIGURACIÓN DEL STMP
        _SMTP.Credentials = New System.Net.NetworkCredential("sdatosnear@gmail.com", "macabeadas")
        _SMTP.Host = "smtp.gmail.com"
        _SMTP.Port = 587
        _SMTP.EnableSsl = True

        ' CONFIGURACION DEL MENSAJE
        _Message.[To].Add(pvstrTo) 'Cuenta de Correo al que se le quiere enviar el e-mail
        _Message.From = New System.Net.Mail.MailAddress(pvstrFrom, pvstrFrom, System.Text.Encoding.UTF8) 'Quien lo envía
        _Message.Subject = pvstrSubject
        _Message.SubjectEncoding = System.Text.Encoding.UTF8 'Codificacion
        _Message.Body = pvstrBody
        _Message.BodyEncoding = System.Text.Encoding.UTF8
        _Message.Priority = System.Net.Mail.MailPriority.Normal
        _Message.IsBodyHtml = False

        'ENVIO
        Try
            _SMTP.Send(_Message)
        Catch ex As System.Net.Mail.SmtpException
            Call GlobalErrorHandler(ex, "Enviando mail")
        End Try

    End Sub

End Class
