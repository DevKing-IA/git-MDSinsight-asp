<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%

ServiceTicketID = Request.Form("txtTicketNumber")
CustID = GetServiceTicketCust(ServiceTicketID)
CustomerName = GetCustNameByCustNum(CustID)
UserNo = Session("UserNo")
UserName = GetUserDisplayNameByUserNo(UserNo)

SourceTab = Request.Form("txtReturnTab")
NewServiceTicketNote = Request.Form("txtNewServiceNote")

NewServiceTicketNote = Replace(NewServiceTicketNote,"'","''")

		
Description = "The service ticket note " & NewServiceTicketNote & ", was created by " & UserName & " on " & FormatDateTime(Now(),2) & " for ticket # from the web app"
Description = Description & ServiceTicketID & ", for customer " & CustomerName & "(" & CustID & ")."

CreateAuditLogEntry "Service ticket note created ",GetTerm("Service"),"Minor",0,Description

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

SQL = "INSERT INTO FS_ServiceMemosNotes (ServiceTicketID, EnteredByUserNo, Note) "
SQL = SQL & "VALUES ('" & ServiceTicketID & "'," & Session("Userno") & ",'" & NewServiceTicketNote & "')"

set rs = cnn8.Execute(SQL)

Set rs = Nothing
Set Cnn8 = Nothing

'********************************************************
'CODE HERE TO MARK NOTE AS BEING READ
 Call MarkNoteNewForUserServiceTicket(ServiceTicketID)
'********************************************************

If SourceTab = "unacknowledged" Then
	Response.redirect("main_UnacknowledgedTickets.asp")
ElseIf SourceTab = "open" Then
	Response.redirect("main_OpenTickets.asp")
Else
	Response.redirect("main_ClosedRedoTickets.asp")
End If

	
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->




