Attribute VB_Name = "SMTP"
Option Compare Database
Option Explicit

Sub mSMTP_SendBuchungsschein()

End Sub

Sub mSMTP_SendMail(sFrom As String, sTo As String, sSubject As String, sBody As String, sAttachmentFile1 As String)
    Dim eMail As Object

    On Error GoTo SendMailErr

    Set eMail = CreateObject("CDO.Message")
    With eMail
        .FROM = sFrom
        .To = sTo
        .Subject = sSubject
        .TextBody = sBody
        If sAttachmentFile1 <> "" Then .AddAttachment sAttachmentFile1
        With .Configuration.Fields
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.mybizcloud.at"
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .Update
        End With
        .Send
    End With

SendMailExit:
    Exit Sub

SendMailErr:
    MsgBox Err.Description, , "mSMTP_SendMail"
    Resume SendMailExit
End Sub

Sub mSMTP_testsend()
    Dim eMail As Object

    Set eMail = CreateObject("CDO.Message")
    With eMail
        .FROM = "releit@gdg-kmsfb.at"
        .To = "office@seus.at"
        .Subject = "Betreff"
        .TextBody = "Mail Text2......"
        '.AddAttachment "c:\FEWO\pdf\buchungsschein1234.pdf"
        With .Configuration.Fields
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.100.248.51"
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            '.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "dialog2840@dsl.at"
            '.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "trx99gtw"
            '.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .Update
        End With
        .Send
    End With

End Sub
