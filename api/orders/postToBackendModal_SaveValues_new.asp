<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs_Orders.asp"-->
<!--#include file="../../inc/InsightFuncs_API.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
OrderID = Request.Form("txtOrderID")
OKtoPost = True

If InternalRecordIdentifier = "" Then OKtoPost = False

	If OKtoPost = True Then
	
			Call rePostOrderToBackend (OrderID,"UPSERT")
	End If

End If

Response.Redirect("main.asp")

'response.end
''Write audit trail for dispatch
''*******************************
'Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
'CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 
'


%>

 
