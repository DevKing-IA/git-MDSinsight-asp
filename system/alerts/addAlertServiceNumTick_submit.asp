<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<%

AlertName = Request.Form("txtAlertName")
ReferenceField = Request.Form("selRefField")
ReferenceValue = Request.Form("selrefVal")
NumberOfTickets = Request.Form("selNumTickets")
NumberOfDays = Request.Form("selNumDays")
SendAlertTo = Request.Form("selSendTo")
AdditionalEmails = Request.Form("txtEmails")
AlertEmailVerbiage = Request.Form("txtVerbiage")
Enabled = Request.Form("chkEnabled")

If Enabled = "on" then Enabled =1 Else Enabled = 0



SQL = "INSERT INTO SC_Alerts (AlertType,AlertName,ReferenceField ,ReferenceValue ,NumberOfTickets , "
SQL = SQL & "NumberOfDays ,SendAlertTo ,Enabled ,AdditionalEmails ,AlertEmailVerbiage,CreatedByUserNo)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'ServiceNumTick'"
SQL = SQL & ",'"  & AlertName & "'"
SQL = SQL & ",'"  & ReferenceField & "'"
SQL = SQL & ",'"  & ReferenceValue & "'"
SQL = SQL & ","  & NumberOfTickets 
SQL = SQL & ","  & NumberOfDays 
SQL = SQL & ",'"  & SendAlertTo & "'"
SQL = SQL & ","  & Enabled 
SQL = SQL & ",'" & AdditionalEmails & "'"	
SQL = SQL & ",'" & AlertEmailVerbiage & "'"	
SQL = SQL & "," & Session("UserNo") & ")"

	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the alert: " & AlertName
CreateAuditLogEntry "Alert Added","Alert Added","Major",0,Description


Response.Redirect("main.asp#ServiceNumTicks")

%>















