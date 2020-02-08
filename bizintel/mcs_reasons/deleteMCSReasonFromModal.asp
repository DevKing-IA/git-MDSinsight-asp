<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->
<%
MCSReasonNoToReplace = Request.Form("txtMCSReasonNoToReplace")
MCSReasonNoReplaceWith = Request.Form("selDeleteMCSReasonFromModal")

If MCSReasonNoToReplace <> "" AND MCSReasonNoReplaceWith <> "" Then
	
	SQLDelete = "SELECT InternalRecordIdentifier FROM BI_MCSReasons WHERE InternalRecordIdentifier = " & MCSReasonNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The MCS reason for this action was changed from ''" & GetMCSReasonByReasonNum(MCSReasonNoToReplace) & "'' to ''" & GetMCSReasonByReasonNum(MCSReasonNoReplaceWith) & "'' to allow for the deletion of ''" & GetMCSReasonByReasonNum(MCSReasonNoToReplace) & "''"
	ReasonDescription = GetMCSReasonByReasonNum(MCSReasonNoToReplace) ' For audit trail below
	
	If NOT rsDelete.EOF Then
		Do
			CreateAuditLogEntry GetTerm("Business Intelligence") & " MCS reason changed",GetTerm("Business Intelligence") & " MCS reason changed","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.EOF
	End If
	rsDelete.Close
	
	'Now replace all BI_MCSActions records with the new MCS Reason number
	
	SQLDelete = "UPDATE BI_MCSActions SET MCSReasonIntRecID = " & MCSReasonNoReplaceWith & " WHERE MCSReasonIntRecID = " & MCSReasonNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion from BI_MCSReason
	
	SQLDelete = "DELETE FROM BI_MCSReasons WHERE InternalRecordIdentifier = " & MCSReasonNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Business Intelligence") & " MCS reason named " & ReasonDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Business Intelligence") & " reason deleted",GetTerm("Business Intelligence") & " reason deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>