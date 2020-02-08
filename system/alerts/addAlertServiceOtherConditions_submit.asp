<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<%
AlertName = Request.Form("txtAlertName")
Enabled = Request.Form("chkEnabled")
Condition = Request.Form("selCond")
Emailto = Request.Form("selEmailto") 
AdditionalEmails = Request.Form("txtaAdditionalEmails")
VerbiageEmail = Request.Form("txtaVerbiageEmail")
Textto = Request.Form("selTextto")
AdditionalTexts = Request.Form("txtaAdditionalTexts")
TextVerbiage = Request.Form("txtAlertTextVerbiage") 
NotificationType = Request.Form("optNotificationType")
PublicOrPrivate = Request.Form("optPublicOrPrivate")

If AdditionalEmails <> "" Then
	AdditionalEmails = Trim(AdditionalEmails)
	AdditionalEmails = Replace(AdditionalEmails,",",";") ' Common for the user to type , instead of ; So we fix it
	If Right(AdditionalEmails,1)=";" Then AdditionalEmails = Left(AdditionalEmails,Len(AdditionalEmails)-1)
End If

If AdditionalTexts <> "" Then
	AdditionalTexts = Trim(AdditionalTexts)
	AdditionalTexts = Replace(AdditionalTexts,",",";") ' Common for the user to type , instead of ; So we fix it
	If Right(AdditionalTexts,1)=";" Then AdditionalTexts = Left(AdditionalTexts,Len(AdditionalTexts)-1)
End If

If Enabled = "on" then Enabled = vbTrue Else Enabled = vbFalse

SQL = "INSERT INTO SC_Alerts (AlertType,AlertName,Condition,EmailToUserNos, "
SQL = SQL & "AdditionalEmails,EmailVerbiage,Enabled ,TextToUserNos,AdditionalText,TextVerbiage,NotificationType,PublicOrPrivate,CreatedByUserNo)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'ServiceOtherConditions'"
SQL = SQL & ",'" & AlertName & "'"
SQL = SQL & ",'" & Condition & "'"
SQL = SQL & ",'" & Emailto & "'"
SQL = SQL & ",'" & AdditionalEmails & "'"
SQL = SQL & ",'" & VerbiageEmail & "'"
SQL = SQL & ","  & Enabled 
SQL = SQL & ",'" & Textto & "'"	
SQL = SQL & ",'" & AdditionalTexts & "'"	
SQL = SQL & ",'" & TextVerbiage & "'"
SQL = SQL & ",'" & NotificationType & "'"
SQL = SQL & ",'" & PublicOrPrivate & "'"
SQL = SQL & "," & Session("UserNo") & ")"


'Response.Write(SQL)
'Response.End
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the alert: " & AlertName
CreateAuditLogEntry "Alert Added","Alert Added","Major",0,Description


Response.Redirect("main.asp#ServiceOtherConditions")
%>















