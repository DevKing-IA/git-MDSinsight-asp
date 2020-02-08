<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/mail.asp"-->

<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

UserNo = Request.QueryString("u")
UserNo = Hacker_Filter1(UserNo)
UserNo = Hacker_Filter2(UserNo)

ClientID = Request.QueryString("c")
ClientID = Hacker_Filter1(ClientID)
ClientID = Hacker_Filter2(ClientID)

If UserNo <> "" Then


	SQL1 = "SELECT * FROM tblServerInfo where clientKey='"& ClientID &"'"
		
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	cnn1.open (InsightCnnString)
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.CursorLocation = 3 
	Set rs1 = cnn1.Execute(SQL1)
		
	If not rs1.EOF Then
		userQuickLoginURL = rs1("QuickLoginURL")
	End If
		
	set rs1 = Nothing
	cnn1.close
	set cnn1 = Nothing

	
	Set ConnectionUsers= Server.CreateObject("ADODB.Connection")
	Set rsUsers = Server.CreateObject("ADODB.Recordset")
	ConnectionUsers.Open Session("ClientCnnString")

	'declare the SQL statement that will query the database
	SQL = "SELECT * FROM tblUsers WHERE userNo= " & userNo

	'Open the recordset object executing the SQL statement and return records
	Set rsUsers = ConnectionUsers.Execute(SQL)

	'If there is no record with the entered username, close connection
	'and go back to login with QueryString
	If rsUsers.EOF then
		rsUsers.close
		ConnectionUsers.close
		set rsUsers =nothing
		set ConnectionUsers=nothing
		%><div><br><font color='red'>No email address found for user number <%= userNo %>.</font></div><%
	Else		
		userEmail = rsUsers("userEmail")
		userFirstName = rsUsers("userFirstName")
		userLastName = rsUsers("userLastName")
		userDisplayName = rsUsers("userDisplayName")
		
		Session("QuickLoginEmailSentTo") = userEmail

		%><!--#include file="../../emails/user_quick_login_credentials.asp"--><%
		
		SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, "System", "Quick Login Credentials"

		Description = "Quick Login URL link email sent to " & userEmail & " for user " & userFirstName & " " & userLastName
	
		CreateAuditLogEntry "Quick Login Credentials Emailed","Quick Login Credentials Emailed","Minor",0,Description 

		%><div><br><font color='green'>Email sent. Please have the user check their email for their quick login credentials.</font></div><%

		ConnectionUsers.close	
		
		Response.Redirect ("edituser.asp?uno=" & UserNo)
	
	End If		

Else

	%><div><br>Unable to send email, could not parse querystring for userno.</div>
	<%
	
End If

%><!--#include file="../../inc/footer-main.asp"-->