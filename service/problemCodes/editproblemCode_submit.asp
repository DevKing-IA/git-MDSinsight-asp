<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM FS_ProblemCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnnproblemCodes = Server.CreateObject("ADODB.Connection")
cnnproblemCodes.open (Session("ClientCnnString"))
Set rsproblemCodes = Server.CreateObject("ADODB.Recordset")
rsproblemCodes.CursorLocation = 3 
Set rsproblemCodes = cnnproblemCodes.Execute(SQL)
	
If not rsproblemCodes.EOF Then
	Orig_ProblemDescription = rsproblemCodes("ProblemDescription")
	Orig_ShowOnWebsite = rsproblemCodes("ShowOnWebsite")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

ProblemDescription = Request.Form("txtProblemDescription")
ShowOnWeb = Request.Form("selShowOnWeb")

SQL = "UPDATE FS_ProblemCodes SET "
SQL = SQL &  "ProblemDescription = '" & ProblemDescription & "' "
SQL = SQL &  ", ShowOnWebsite = '" & ShowOnWeb & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set rsproblemCodes = cnnproblemCodes.Execute(SQL)
set rsproblemCodes = Nothing


Description = ""
If Orig_ProblemDescription  <> ProblemDescription  Then
	Description = Description & "Service module problem code changed from " & Orig_ProblemDescription  & " to " & ProblemDescription  
End If

CreateAuditLogEntry "Service module problem code edited","Service module problem code edited","Minor",0,Description

Response.Redirect("main.asp")

%>















