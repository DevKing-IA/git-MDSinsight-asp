<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM FS_SymptomCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnnsymptomCodes = Server.CreateObject("ADODB.Connection")
cnnsymptomCodes.open (Session("ClientCnnString"))
Set rssymptomCodes = Server.CreateObject("ADODB.Recordset")
rssymptomCodes.CursorLocation = 3 
Set rssymptomCodes = cnnsymptomCodes.Execute(SQL)
	
If not rssymptomCodes.EOF Then
	Orig_SymptomDescription = rssymptomCodes("SymptomDescription")
	Orig_ShowOnWebsite = rssymptomCodes("ShowOnWebsite")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SymptomDescription = Request.Form("txtSymptomDescription")
ShowOnWeb = Request.Form("selShowOnWeb")

SQL = "UPDATE FS_SymptomCodes SET "
SQL = SQL &  "SymptomDescription = '" & SymptomDescription & "' "
SQL = SQL &  ", ShowOnWebsite = '" & ShowOnWeb & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set rssymptomCodes = cnnsymptomCodes.Execute(SQL)
set rssymptomCodes = Nothing


Description = ""
If Orig_SymptomDescription  <> SymptomDescription  Then
	Description = Description & "Service module symptom code changed from " & Orig_SymptomDescription  & " to " & SymptomDescription  
End If

CreateAuditLogEntry "Service module symptom code edited","Service module symptom code edited","Minor",0,Description

Response.Redirect("main.asp")

%>















