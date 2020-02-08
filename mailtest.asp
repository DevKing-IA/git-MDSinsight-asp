<%
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress "rich@ocsaccess.com" , "rich@ocsaccess.com"
	Mailer.From = emailFrom
	Mailer.FromName = "mailsender@dev.mdsinsight.com"
	Mailer.Subject = "test"
	Mailer.IsHTML = True
	Mailer.Body = "test"

	strErr = ""
	bSuccess = False
	On Error Resume Next ' catch errors
	Mailer.Send	' send message
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	else
	  bSuccess = True
	End If
%>