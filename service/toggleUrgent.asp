<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%
MemoNum = Request.Form("memnum")

Set cnnToggle = Server.CreateObject("ADODB.Connection")
cnnToggle.open (Session("ClientCnnString"))
Set rsToggle = Server.CreateObject("ADODB.Recordset")
rsToggle.CursorLocation = 3 

SQL = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MemoNum  & "'"
Set rsToggle = cnnToggle.Execute(SQL)

If Not rsToggle.Eof Then
	'All records should be the same so just check the 1st one
	If rsToggle("Urgent") <> 1 Then 
		Urgent = 1
		CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to urgent"
	Else
		Urgent = 0
		CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to not urgent"		
	End If	
	SQL = "UPDATE FS_ServiceMemos Set Urgent = " & Urgent & " WHERE MemoNumber= '" & MemoNum & "'"
	Set rsToggle = cnnToggle.Execute(SQL)
End If

set rsToggle = Nothing
cnnToggle.Close
Set cnnToggle = Nothing

%>
