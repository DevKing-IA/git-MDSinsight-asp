<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<%
AlertName = Request.Form("txtAlertName")
Enabled = Request.Form("chkEnabled")
Condition = Request.Form("selCond")
Minutes = Request.Form("selMinutes")
Emailto = Request.Form("selEmailto") 
AdditionalEmails = Request.Form("txtaAdditionalEmails")
VerbiageEmail = Request.Form("txtaVerbiageEmail")
Textto = Request.Form("selTextto")
AdditionalTexts = Request.Form("txtaAdditionalTexts")
TextVerbiage = Request.Form("txtAlertTextVerbiage") 
IncludeLog = Request.Form("chkLog")
LimitMinutes = Request.Form("selLimitMinutes")
LimitMaxTimes = Request.Form("selLimitMaxTimes")
NotificationType = Request.Form("optNotificationType")
PublicOrPrivate = Request.Form("optPublicOrPrivate")

'Response.Write("LimitMinutes " & LimitMinutes & "<br>")
'Response.Write("Request.Form(selLimitMinutes) " & Request.Form("selLimitMinutes") & "<br>")
'Response.Write("LimitMaxTimes " & LimitMaxTimes & "<br>")
'Response.Write("Request.Form(selLimitMaxTimes) " & Request.Form("selLimitMaxTimes") & "<br>")
'Response.End

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

If IncludeLog = "on" then IncludeLog = vbTrue Else IncludeLog = vbFalse
If Enabled = "on" then Enabled = vbTrue Else Enabled = vbFalse
If LimitMinutes = "" Then LimitMinutes = 60
If LimitMaxTimes = "" Then LimitMaxTimes = 1


SQL = "INSERT INTO SC_Alerts (AlertType,AlertName,Condition,NBMinutes,EmailToUserNos, "
SQL = SQL & "AdditionalEmails,EmailVerbiage,Enabled ,TextToUserNos,AdditionalText,TextVerbiage,NBIncludeLog,NBLimitMiniutes,NBLimitMaxTimes,NotificationType,PublicOrPrivate,CreatedByUserNo)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'NightBatch'"
SQL = SQL & ",'" & AlertName & "'"
SQL = SQL & ",'" & Condition & "'"
SQL = SQL & ","  & Minutes & ""
SQL = SQL & ",'" & Emailto & "'"
SQL = SQL & ",'" & AdditionalEmails & "'"
SQL = SQL & ",'" & VerbiageEmail & "'"
SQL = SQL & ","  & Enabled 
SQL = SQL & ",'" & Textto & "'"	
SQL = SQL & ",'" & AdditionalTexts & "'"	
SQL = SQL & ",'" & TextVerbiage & "'"
SQL = SQL & "," & IncludeLog 
SQL = SQL & "," & LimitMinutes 
SQL = SQL & "," & LimitMaxTimes
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


Response.Redirect("main.asp#NightBatchAlerts")
%>















