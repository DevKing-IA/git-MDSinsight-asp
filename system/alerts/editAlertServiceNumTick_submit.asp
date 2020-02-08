<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<%
InternalAlertRecNumber = Request.Form("txtInternalAlertRecNumber")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM SC_Alerts where InternalAlertRecNumber = " & InternalAlertRecNumber 
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_AlertName = rs("AlertName")
	Orig_ReferenceField = rs("ReferenceField")
	Orig_ReferenceValue = rs("ReferenceValue")
	Orig_NumberOfTickets = rs("NumberOfTickets")
	Orig_NumberOfDays = rs("NumberOfDays")
	Orig_SendAlertTo = rs("SendAlertTo")
	Orig_AdditionalEmails = rs("AdditionalEmails")
	Orig_AlertEmailVerbiage = rs("AlertEmailVerbiage")
	Orig_Enabled = rs("Enabled")
	Orig_InternalAlertRecNumber = rs("InternalAlertRecNumber")
	Orig_LimitMinutes = rs("NBLimitMiniutes")
	Orig_LimitMaxTimes = rs("NBLimitMaxTimes")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

AlertName = Request.Form("txtAlertName")
ReferenceField = Request.Form("selRefField")
ReferenceValue = Request.Form("selrefVal")
NumberOfTickets = Request.Form("selNumTickets")
NumberOfDays = Request.Form("selNumDays")
SendAlertTo = Request.Form("selSendTo")
AdditionalEmails = Request.Form("txtEmails")
AlertEmailVerbiage = Request.Form ("txtVerbiage")
Enabled = Request.Form("chkEnabled")
InternalAlertRecNumber = Request.Form("txtInternalAlertRecNumber")



If Enabled = "on" then Enabled =1 Else Enabled = 0

SQL = "UPDATE SC_Alerts SET "
SQL = SQL &  "AlertName = '" & AlertName & "',"
SQL = SQL &  "ReferenceField = '" & ReferenceField & "',"
SQL = SQL &  "ReferenceValue = '" & ReferenceValue & "',"
SQL = SQL &  "NumberOfTickets = " & NumberOfTickets & ","
SQL = SQL &  "NumberOfDays = " & NumberOfDays & ","
SQL = SQL &  "SendAlertTo = '" & SendAlertTo & "',"
SQL = SQL &  "AdditionalEmails = '" & AdditionalEmails & "',"
SQL = SQL &  "AlertEmailVerbiage = '" & AlertEmailVerbiage & "',"
SQL = SQL &  "Enabled = " & Enabled 
SQL = SQL &  " WHERE InternalAlertRecNumber = " & InternalAlertRecNumber 
	
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_AlertName <> AlertName Then
	Description = Description & "Alert name changed from " & Orig_AlertName & " to " & AlertName 
End If
If Orig_ReferenceField <> ReferenceField Then
	Description = Description & "  Reference field changed from " & Orig_ReferenceField & " to " & ReferenceField 
End If
If Orig_ReferenceValue <> ReferenceValue Then
	Description = Description & "  Reference value changed from " & Orig_ReferenceValue & " to " & ReferenceValue 
End If
If Orig_NumberOfTickets <> NumberOfTickets Then
	Description = Description & "  Number of tickets changed from " & Orig_NumberOfTickets & " to " & NumberOfTickets 
End If
If Orig_NumberOfDays <> NumberOfDays Then
	Description = Description & "  Number of days changed from " & Orig_NumberOfDays & " to " & NumberOfDays 
End If
If Orig_SendAlertTo <> SendAlertTo Then
	Description = Description & "  Send alerts to changed from " & Orig_SendAlertTo & " to " & SendAlertTo 
End If
If Orig_AdditionalEmails <> AdditionalEmails Then
	Description = Description & "  Send additional alerts to changed from " & Orig_AdditionalEmails & " to " & AdditionalEmails 
End If
If Orig_AlertEmailVerbiage <> AlertEmailVerbiage Then
	Description = Description & "  Alert email verbiage changed from " & Orig_AlertEmailVerbiage & " to " & AlertEmailVerbiage 
End If
If Orig_Enabled <> Enabled Then
	Description = Description & "  Enabled changed from " & Orig_Enabled & " to " & Enabled 
End If


CreateAuditLogEntry "Alert Edited","Alert Edited","Major",0,Description

Response.Redirect("main.asp#ServiceNumTicks")

%>















