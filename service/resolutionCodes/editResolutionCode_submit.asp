<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM FS_ResolutionCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnnresolutionCodes = Server.CreateObject("ADODB.Connection")
cnnresolutionCodes.open (Session("ClientCnnString"))
Set rsresolutionCodes = Server.CreateObject("ADODB.Recordset")
rsresolutionCodes.CursorLocation = 3 
Set rsresolutionCodes = cnnresolutionCodes.Execute(SQL)
	
If not rsresolutionCodes.EOF Then
	Orig_ResolutionDescription = rsresolutionCodes("ResolutionDescription")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

ResolutionDescription = Request.Form("txtResolutionDescription")
ShowOnWeb = Request.Form("selShowOnWeb")

SQL = "UPDATE FS_ResolutionCodes SET "
SQL = SQL &  "ResolutionDescription = '" & ResolutionDescription & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set rsresolutionCodes = cnnresolutionCodes.Execute(SQL)
set rsresolutionCodes = Nothing


Description = ""
If Orig_ResolutionDescription  <> ResolutionDescription  Then
	Description = Description & "Service module resolution code changed from " & Orig_ResolutionDescription  & " to " & ResolutionDescription  
End If

CreateAuditLogEntry "Service module resolution code edited","Service module resolution code edited","Minor",0,Description

Response.Redirect("main.asp")

%>















