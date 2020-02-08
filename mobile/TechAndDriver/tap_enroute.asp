<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/settings.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way

ServiceTicketNumber = Request.QueryString("t")
UserNumber = Request.QueryString("u")
CustNum = Request.QueryString("c")

'Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
'Response.Write("UserNumber :" & UserNumber & "<br>")
'Response.Write("CustNum :" & CustNum & "<br>")
'Response.End

If ServiceTicketNumber <> "" AND UserNumber <> "" AND CustNum <> "" Then
	
	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,Urgent,OriginalDispatchDateTime)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'En Route'"
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & UserNumber 
	SQLDispatch = SQLDispatch & ","  & GetServiceTicketDispatchedTech(ServiceTicketNumber )
	If TicketIsUrgent(ServiceTicketNumber) Then
		SQLDispatch = SQLDispatch & ",1" 'Urgent
	Else
		SQLDispatch = SQLDispatch & ",0" 'Not Urgent
	End If
	SQLDispatch = SQLDispatch & ", '" & TicketOriginalDispatchDateTime(ServiceTicketNumber) & "')"

	Set cnnDispatch = Server.CreateObject("ADODB.Connection")
	cnnDispatch.open (Session("ClientCnnString"))
	Set rsDispatch = Server.CreateObject("ADODB.Recordset")
	Set rsDispatch = cnnDispatch.Execute(SQLDispatch)

	
	'Write audit trail for dispatch
	'*******************************
	Description = GetUserDisplayNameByUserNo(UserNumber) & " changed ticket stage to En Route for service ticket number " & ServiceTicketNumber & " at " & NOW()
	CreateAuditLogEntry "Service Ticket System","En ROute","Minor",0,Description 
	
	Set rsDispatch = Nothing
	cnnDispatch.Close
	Set cnnDispatch = Nothing

End If

Response.redirect("main.asp")
%>

 
