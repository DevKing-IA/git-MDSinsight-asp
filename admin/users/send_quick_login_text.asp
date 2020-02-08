<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

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

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV.") <> 0 AND Left(ucase(userQuickLoginURL),11) <> "HTTP://DEV." Then 
	
		'get rid of http://
		userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)
	
		'Strip the URL first part
		For x = 1 to len(userQuickLoginURL)
			If Mid(userQuickLoginURL,x,1)="." Then
				userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
				Exit For
			End If
		Next 
	
		userQuickLoginURL = "http://dev." & userQuickLoginURL 
	
	End If
	
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV2.") <> 0 AND Left(ucase(userQuickLoginURL),12) <> "HTTP://DEV2." Then 
	
		'get rid of http://
		userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)
	
		'Strip the URL first part
		For x = 1 to len(userQuickLoginURL)
			If Mid(userQuickLoginURL,x,1)="." Then
				userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
				Exit For
			End If
		Next 
	
		userQuickLoginURL = "http://dev2." & userQuickLoginURL 
	
	End If
		
	set rs1 = Nothing
	cnn1.close
	set cnn1 = Nothing

	
	SQL2 = "SELECT * FROM tblUsers WHERE userNo="& UserNo 
	
	Set cnn2 = Server.CreateObject("ADODB.Connection")
	cnn2.open (Session("ClientCnnString"))
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.CursorLocation = 3 
	Set rs2 = cnn2.Execute(SQL2)

	If not rs2.eof then
	
		QuickLoginURL = userQuickLoginURL & "?u=" & userNo & "%26c=" & ClientID

		URL = BaseURL & "inc/sendtext.php"
		
		QString = "?n=" & Replace(rs2("userCellNumber"),"-","")
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		QString = QString & "&t=MDSInsight"
		QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "admin/users/edituser.asp?uno=" & UserNo)
		QString = QString & "&m=Quick Login Info" & vbCRLF
		QString = QString &  "Password: " & rs2("userPassword") & vbCRLF
		QString = QString &  "Quick Link: " & QuickLoginURL
		QString = QString &  "&cty=" & GetCompanyCountry()
		QString = Replace(Qstring," ", "%20")


		Session("QuickLoginTextSentTo") = rs2("userCellNumber")
		
		CreateAuditLogEntry "Quick Login Credentials Texted","Quick Login Credentials Texted","Minor",0,"Quick login URL text sent to " & rs2("userCellNumber") & " for user " & rs2("userFirstName") & " " & rs2("userLastName")

		set rs2 = Nothing
		set cnn2  = Nothing
						
		Response.Redirect (URL & Qstring)
				
	Else
	
		%><div><br />
		Unable to send text message, user not found: <%= userEmail %>.
		</div>
		<%
		
	End If
	
	Else
	
		%><div><br />
		Unable to send text message, could not parse querystring for userno.
		</div>
		<%
	
End If
%><!--#include file="../../inc/footer-main.asp"-->