Imports System.Net.Mail
Imports System.Configuration
Public Class clsMail
    Dim sFrom As String = ConfigurationSettings.AppSettings("From")
    Dim sBCC As String = ConfigurationSettings.AppSettings("BCC")
    Dim sCC As String = ConfigurationSettings.AppSettings("CC")
    Dim smtpClient As String = ConfigurationSettings.AppSettings("MailServ")

    Public Sub EmailResults(ByVal sSubject As String, ByVal sBody As String, ByVal sTo As String)
        Dim Message As New MailMessage(sFrom, sTo)
        Try
            With Message
                .CC.Add(sCC)
                .Subject = sSubject
                .Body = sBody
            End With
            Dim Client As New SmtpClient(smtpClient)
            Client.UseDefaultCredentials = True
            Client.Send(Message)
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub EmailFiles(ByVal sSubject As String, ByVal sBody As String, ByVal sTo As String, ByVal sAttach As String)
        Dim Message As New MailMessage(sFrom, sTo)
        Dim attach As New Net.Mail.Attachment(sAttach)
        Try
            With Message
                .Attachments.Add(attach)
                .CC.Add(sCC)
                .Subject = sSubject
                .Body = sBody
            End With
            Dim Client As New SmtpClient(smtpClient)
            Client.UseDefaultCredentials = True
            Client.Send(Message)
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

End Class