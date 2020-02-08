<!--#include file="../../inc/header.asp"-->


<%

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
	
	
		URL = BaseURL & "inc/sendtext.php"
		QString = "?n=" & Replace(rs("userCellNumber"),"-","")
		
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=MDSInsight"
		QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "admin/users/main.asp#" & ActiveTab)
		QString = QString & "&m=Login Info" & vbCRLF
		QString = QString &  "Email: " & rs("userEmail") & vbCRLF
		QString = QString &  "Password: " & rs("userPassword") & vbCRLF
		QString = QString &  "Client Key: " &  MUV_Read("ClientID") & vbCRLF
		QString = QString &  "Go to: " & BaseURL 
		QString = QString &  "&cty=" & GetCompanyCountry()
	
		
		QString = Replace(Qstring," ", "%20")

		Session("LoginTextSentTo") = rs("userCellNumber")
		
		CreateAuditLogEntry "Credentials Sent","Credentials Sent","Minor",0,"Login credential/welcome text sent to " & rs("userCellNumber") & " for user " & rs("userFirstName") & " " & rs("userLastName")
				
		Response.Redirect (URL & Qstring)
		
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