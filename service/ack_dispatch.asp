<!--#include file="../inc/InsightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/settings.asp"-->

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

If ServiceTicketNumber = "" or UserNumber= "" or CustNum = "" Then Response.redirect(baseURL)


SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQLDispatch = SQLDispatch & " SubmissionDateTime, USerNoSubmittingRecord)"
SQLDispatch = SQLDispatch &  " VALUES (" 
SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
SQLDispatch = SQLDispatch & ",'Dispatch Acknowledged'"
SQLDispatch = SQLDispatch & ",getdate() "
SQLDispatch = SQLDispatch & ","  & UserNumber & ")"

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)

'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(UserNumber) & " acknowldged dispatch notification for service ticket number " & ServiceTicketNumber & " at " & NOW()
CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 

Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing


Response.Write("<font color='blue' size='6'><center><br><br><br><br>Your acknowledgement has been recorded.<br><br><br><br> Thank You.</font></center>")
Response.End
%>

 
