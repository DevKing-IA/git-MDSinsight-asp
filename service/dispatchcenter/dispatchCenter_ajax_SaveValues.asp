<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

SendEmail = Request.Form("chkSendEmail")


SendText = Request("chkSendText")
UserToDispatch = Request("selFieldTech")
ServiceTicketNumber = Request("txtServiceTicketNumber")

CustNum = GetServiceTicketCust(ServiceTicketNumber)

'CustNum = Request("txtAccountNumber") ' old code



If UserToDispatch <> 0 then ' If 0, it ia an Undispatch
	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'Dispatched'"
	SQLDispatch = SQLDispatch & ","  & UserToDispatch 
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & Session("UserNo")
	SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
	SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
	SQLDispatch = SQLDispatch & ", getDate()"	
	SQLDispatch = SQLDispatch & ",'" &  GetUserDisplayNameByUserNo(UserToDispatch) & " has been dispatched. This ticket was reassigned via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & "')"	
Else
	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'Received'"
	SQLDispatch = SQLDispatch & ","  & UserToDispatch 
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & Session("UserNo")
	SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
	SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
	SQLDispatch = SQLDispatch & ", getDate()"	
	SQLDispatch = SQLDispatch & ",'This ticket was changed from " &  GetServiceTicketCurrentStage(ServiceTicketNumber) & " to un-dispatched via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & "')"	
	HeldStatus = GetServiceTicketCurrentStage(ServiceTicketNumber)
	HeldTech = GetServiceTicketDispatchedTech(ServiceTicketNumber)
End If

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)


'Write audit trail for dispatch
'*******************************
If UserToDispatch <> 0 Then
	Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched via the dispatch center to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
Else
	Description = "Service ticket " & ServiceTicketNumber  & " was changed from " &  HeldStatus  & " to un-dispatched via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
End If
CreateAuditLogEntry "Dispatch Center","Dispatch Center","Major",0,Description 

If UserToDispatch <> 0 Then
'Also set dispatched flag for simple dispatch model
	SQLDispatch = "Update FS_ServiceMemos Set Dispatched = CASE WHEN Dispatched = 0 THEN -1 ELSE 0 END Where MemoNumber = '"  & ServiceTicketNumber & "'"
	Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
End If

Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing

retData="{""UserToDispatch"":""" & UserToDispatch & """,""ServiceTicketNumber"":""" & ServiceTicketNumber & """,""UserDisplayName"":""" & GetUserDisplayNameByUserNo(UserToDispatch) & """,""UserEmailAddress"":""" & getUserEmailAddress(UserToDispatch) & """}"
Response.Write(retData)
Response.end


%>

 
