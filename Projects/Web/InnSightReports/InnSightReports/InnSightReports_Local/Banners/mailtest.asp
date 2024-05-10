
<form method="POST" action="mailtest.asp">
  <div align="center">
    <center>
    <table border="1" cellpadding="6" cellspacing="0" width="450" bordercolor="#000000">
      <tr>
        <td width="446" colspan="2" bgcolor="#C0C0C0">
          <p align="center"><font face="Arial" size="5">Email Tester</font></p>
          <p align="center"><font face="Arial" size="2">This will test to see
          what email components are installed on your server.&nbsp; For this
          test to work properly, you must be certain your email address if valid
          and that you have specified a valid email server.</font></td>
      </tr>
      <tr>
        <td width="145" align="right"><font face="Arial" size="2">Email Address:</font></td>
        <td width="301"><input type="text" name="Email" size="40" value="<%=Application("AdministratorEmail")%>"></td>
      </tr>
      <tr>
        <td width="145" align="right"><font face="Arial" size="2">Mail Server:</font></td>
        <td width="301"><input type="text" name="MailServer" size="40" value="<%=Application("MailServer")%>"></td>
      </tr>
      <tr>
        <td width="145" align="right"> </td>
        <td width="301"><input type="submit" value="Perform Test" name="B1"></td>
      </tr>
    </table>
    
<% 
If Request.Form("Email") <> "" And Request.Form("MailServer")<> "" Then
'******* CDONTS Mail ********************************************************
On Error Resume Next
strSMPRecipient=Request.Form("Email")
strBMPMailServer=Request.Form("MailServer")
strBMPSubject="Email Test"

arrRecipients=Split(strSMPRecipient,",")
arrTemp=Split(Application("AdministratorEmail"),",")
strBMPFrom=arrTemp(0)

strBMPMessage="*** Test Succeeded"

	Set Mailer = Server.CreateObject("CDONTS.NewMail")
	Mailer.To = strSMPRecipient
	Mailer.From = strSMPRecipient
	Mailer.Subject = strBMPSubject
	Mailer.Body = strBMPMessage & " for CDONTS"
	Mailer.Send
	Set Mailer = nothing
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "<p>Test Failed for CDONTS<br>"
		Err.Clear
	Else 
		Response.Write "<p><font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for CDONTS</font><br>"
	End If
'******* J Mail ********************************************************
	Set Mailer = Server.CreateObject("JMail.SMTPMail")
	Mailer.ServerAddress = strBMPMailServer
	Mailer.Sender = strBMPFrom
	Mailer.SenderName =arrRecipients(0)
	intCnt=0
	Do While intCnt<=Ubound(arrRecipients)
		Mailer.AddRecipientEx arrRecipients(intCnt), arrRecipients(intCnt)
		intCnt=intCnt+1
	Loop
	Mailer.Subject = strBMPSubject
	Mailer.Body = strBMPMessage & " for J Mail"
	Mailer.Execute
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "Test Failed for J Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for J Mail</font><br>"
	End If
'******* Simple Mail ********************************************************
	Set Mailer = Server.CreateObject("SimpleMail.smtp.1")
	Mailer.OpenConnection strBMPMailServer
	Mailer.SendMail strSMPRecipient, strSMPRecipient, strBMPSubject, strBMPMessage
	Mailer.CloseConnection
	Mailer.OpenConnection strBMPMailServer
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "Test Failed for Simple Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for Simple Mail</font><br>"
	End If
'******* ASP Mail ********************************************************
	Set Mailer = Server.CreateObject("SMTPSVG.Mailer")
	Mailer.RemoteHost = strBMPMailServer
	intCnt=0
	Do While intCnt<=Ubound(arrRecipients)
		Mailer.AddRecipient " ", arrRecipients(intCnt)
		intCnt=intCnt+1
	Loop
	Mailer.FromAddress =strBMPFrom
	Mailer.FromName = strBMPFrom
	Mailer.Subject = strBMPSubject
	Mailer.BodyText = strBMPMessage  & " for ASP Mail"
	Mailer.SendMail
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "Test Failed for ASP Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for ASP Mail</font><br>"
	End If
'******* Persits ASP EMail ********************************************************
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
	Mailer.Body = strBMPMessage  & " for ASP Mail"
	Mailer.Send
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "Test Failed for Persits ASP EMail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for Persits ASP EMail</font><br>"
	End If
'******* AB Mail ********************************************************
	Set Mailer = Server.CreateObject("ABMailer.Mailman")
	Mailer.Clear
	Mailer.SendTo = strSMPRecipient
	Mailer.ReplyTo = strBMPFrom
	Mailer.MailSubject = strBMPSubject
	Mailer.MailDate =""
	Mailer.ServerAddr = strBMPMailServer
	Mailer.MailMessage = strBMPMessage  & " for AB Mail"
	Mailer.SendMail
	Set Mailer = nothing	
	If Err.Number >0 Then
		Response.Write "Test Failed for AB Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for AB Mail</font><br>"
	End If
'******* Bamboo Mail ****************************************************
	Set Mailer = Server.CreateObject("Bamboo.SMTP")
	Mailer.Server = strBMPMailServer
	Mailer.RCPT = strSMPRecipient
	Mailer.From = strBMPFrom
	Mailer.FromName = strBMPFrom
	Mailer.Subject = strBMPSubject
	Mailer.Message = strBMPMessage  & " for Bamboo Mail"
	Mailer.Send
	Set Mailer = nothing	
	If Err.Number >0 Then
		Response.Write "Test Failed for Bamboo Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for Bamboo Mail</font><br>"
	End If
'******* OCX Mail ********************************************************
	Set Mailer = Server.CreateObject("ASPMail.ASPMailCtrl.1")
	Mailer.Send
	Mailer.SendMail strBMPMailServer, strSMPRecipient, strBMPFrom, strBMPSubject, strBMPMessage  & " for OCX Mail"
	Set Mailer = nothing
	If Err.Number >0 Then
		Response.Write "Test Failed for OCX Mail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for OCX Mail</font><br>"
	End If
'******* OCX Mail ********************************************************
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
                                        strBMPMessage  & " for OCX QMail")

	Set Mailer = nothing
		If Err.Number >0 Then
		Response.Write "Test Failed for OCX QMail<br>"
		Err.Clear
	Else 
		Response.Write "<font color=" & chr(34) & "#0000FF" & chr(34) & ">*** Test Succeeded for OCX QMail</font><br>"
	End If
End If

%>
    
    </center>
  </div>
</form>
