<% 
arrTemp=Split(Application("AdministratorEmail"),",")
strBMPFrom=arrTemp(0)
arrRecipients=Split(strSMPRecipient,",")
'******* CDONTS Mail ********************************************************
If Application("MailProgram") = "CDONTS Mail" Then
	Set Mailer = Server.CreateObject("CDONTS.NewMail")
	Mailer.To = strSMPRecipient
	Mailer.From = strSMPRecipient
	Mailer.Subject = strBMPSubject
	Mailer.Body = strBMPMessage
	Mailer.Send
	Set Mailer = nothing
	Set Mailer = nothing
'******* J Mail ********************************************************
ElseIf Application("MailProgram") = "J Mail" Then
	Set Mailer = Server.CreateObject("JMail.SMTPMail")
	Mailer.ServerAddress = strBMPMailServer
	Mailer.Sender = strFromEmail
	Mailer.SenderName =strBMPFrom
	intCnt=0
	Do While intCnt<=Ubound(arrRecipients)
		Mailer.AddRecipientEx arrRecipients(intCnt), arrRecipients(intCnt)
		intCnt=intCnt+1
	Loop
	Mailer.Subject = strBMPSubject
	Mailer.Body = strBMPMessage
	Mailer.Execute
	Set Mailer = nothing
'******* Simple Mail ********************************************************
ElseIf Application("MailProgram") = "Simple Mail" Then
	Set Mailer = Server.CreateObject("SimpleMail.smtp.1")
	Mailer.OpenConnection strBMPMailServer
	Mailer.SendMail strSMPRecipient, strSMPRecipient, strBMPSubject, strBMPMessage
	Mailer.CloseConnection
	Mailer.OpenConnection strBMPMailServer
	Set Mailer = nothing
'******* ASP Mail ********************************************************
ElseIf Application("MailProgram") = "ASP Mail" Then 
	Set Mailer = Server.CreateObject("SMTPSVG.Mailer")
	Mailer.RemoteHost = strBMPMailServer
	intCnt=0
	Do While intCnt<=Ubound(arrRecipients)
		Mailer.AddRecipient " ", arrRecipients(intCnt)
		intCnt=intCnt+1
	Loop
	Mailer.FromAddress = strBMPFrom
	Mailer.FromName = strBMPFrom
	Mailer.Subject = strBMPSubject
	Mailer.BodyText = strBMPMessage
	Mailer.SendMail
	Set Mailer = nothing
'******* ASP Mail ********************************************************
ElseIf Application("MailProgram") = "Persits ASP EMail" Then 
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = strBMPMailServer
	intCnt=0
	Do While intCnt<=Ubound(arrRecipients)
		Mailer.AddAddress arrRecipients(intCnt)," "
		intCnt=intCnt+1
	Loop
	Mailer.From = strBMPFrom
	Mailer.FromName = strBMPFrom
	Mailer.Subject = strBMPSubject
	Mailer.Body = strBMPMessage
	Mailer.Send
	Set Mailer = nothing
'******* AB Mail ********************************************************
ElseIf Application("MailProgram")= "AB Mail" Then
	Set Mailer = Server.CreateObject("ABMailer.Mailman")
	Mailer.Clear
	Mailer.SendTo = strSMPRecipient
	Mailer.ReplyTo = strBMPFrom
	Mailer.MailSubject = strBMPSubject
	Mailer.MailDate =""
	Mailer.ServerAddr = strBMPMailServer
	Mailer.MailMessage = strBMPMessage
	Mailer.SendMail
	Set Mailer = nothing	
'******* Bamboo Mail ****************************************************
ElseIf Application("MailProgram") = "Bamboo Mail" Then
	Set Mailer = Server.CreateObject("Bamboo.SMTP")
	Mailer.Server = strBMPMailServer
	Mailer.RCPT = strSMPRecipient
	Mailer.From = strBMPFrom
	Mailer.FromName = strBMPFrom
	Mailer.Subject = strBMPSubject
	Mailer.Message = strBMPMessage
	Mailer.Send
	Set Mailer = nothing	
'******* OCX Mail ********************************************************
ElseIf Application("MailProgram") = "OCX Mail" Then
	Set Mailer = Server.CreateObject("ASPMail.ASPMailCtrl.1")
	Mailer.Send
	Mailer.SendMail strBMPMailServer, strSMPRecipient, strBMPFrom, strBMPSubject, strBMPMessage
	Set Mailer = nothing
'******* OCX Mail ********************************************************
ElseIf Application("MailProgram") = "OCXQMail" Then
	Set Mailer = Server.CreateObject("OCXQmail.OCXQmailCtrl.1")
	mailServer = strBMPMailServer
	mailer.SendAt(Now)
	result = mailer.Q(mailServer, _
                                        strBMPFrom, _
                                        strBMPFrom, _
                                        priority, _
                                        "", _
                                        strSMPRecipient, _
                                        ccAddressList, _
                                        bccAddressList, _
                                        attachmentList, _
                                        strBMPSubject, _
                                        strBMPMessage)

	Set Mailer = nothing
End If

%>