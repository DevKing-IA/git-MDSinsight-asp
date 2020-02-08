<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/mail.asp"-->



<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

ActiveTab = Request.QueryString("tab")

UserNo = Request.QueryString("userno")
UserNo = Hacker_Filter1(UserNo)
UserNo = Hacker_Filter2(UserNo)

If UserNo <> "" Then
	
	SQL = "SELECT * FROM tblUsers WHERE userNo="& UserNo 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.eof then
	
	
		%> <!--#include file="../../emails/user_login_credentials.asp"--> <%
		
		userEmail = rs("userEmail")
		SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, "System", "Login Credentials"

		Session("LoginEmailSentTo") = userEmail
		
		Description = "Login credential/welcome email sent to " & userEmail & " for user " & rs("userFirstName") & " " & rs("userLastName")
		
		CreateAuditLogEntry "Credentials Sent","Credentials Sent","Minor",0,Description 

		
		Response.Redirect ("main.asp#" & ActiveTab)

	Else
	
		%><div><br />
		Unable to send, user not found: <%= userEmail %>.
		</div>
		<%
		
	End If
	
	Else
	
		%><div><br />
		Unable to send, could not parse querystring for userno.
		</div>
		<%
	
End If

set rs = Nothing
set cnn8  = Nothing
%><!--#include file="../../inc/footer-main.asp"-->