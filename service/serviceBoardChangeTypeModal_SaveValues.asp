<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%
NewType = Request.Form("optTicketType")
ServiceTicketNumber = Request.Form("txtServiceTicketNumber")
CustNum = Request.Form("txtAccountNumber")

Response.Write("NewType :" & NewType & "<br>")
Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
Response.Write("BaseURL :" & BaseURL & "<br>")
Response.Write("CustNum :" & CustNum & "<br>")
'Response.End

If NewType ="Service Ticket" Then
	SQLChangeType = "UPDATE FS_ServiceMemos SET FilterChange = 0 ,PMCall = 0 WHERE MemoNumber = '" & ServiceTicketNumber & "'"
End If	

If NewType ="Filter Change" Then
	SQLChangeType = "UPDATE FS_ServiceMemos SET FilterChange = 1 ,PMCall = 0 WHERE MemoNumber = '" & ServiceTicketNumber & "'"
End If	

If NewType ="PM Call" Then
	SQLChangeType = "UPDATE FS_ServiceMemos SET FilterChange = 0 ,PMCall = 1 WHERE MemoNumber = '" & ServiceTicketNumber & "'"
End If	

Response.Write(SQLChangeType & "<br>")

Set cnnChangeType = Server.CreateObject("ADODB.Connection")
cnnChangeType.open (Session("ClientCnnString"))
Set rsChangeType = Server.CreateObject("ADODB.Recordset")
Set rsChangeType = cnnChangeType.Execute(SQLChangeType)


'Write audit trail for ticket type chnage
'*******************************
'If UserToDispatch <> 0 Then
'	Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched via the dispatch center to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
'Else
'	Description = "Service ticket " & ServiceTicketNumber  & " was changed from " &  HeldStatus  & " to un-dispatched via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
'End If
'CreateAuditLogEntry "Dispatch Center","Dispatch Center","Major",0,Description 

Set rsChangeType = Nothing
cnnChangeType.Close
Set cnnChangeType = Nothing


Response.Redirect(BaseURL & "service/main.asp")
%>

 
