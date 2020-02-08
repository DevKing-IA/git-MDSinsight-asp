<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
txtProspectCurrentComments = Request.Form("txtProspectCurrentComments")
txtProspectEditComments = Request.Form("txtProspectEditComments")

If txtProspectEditComments = "" Then txtProspectEditComments = "COMMENTS REMOVED"
If txtProspectCurrentComments = "" Then txtProspectCurrentComments = "NONE ENTERED"

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	

If txtInternalRecordIdentifier <> "" Then
	
	'Update prospect comments

	Set cnnProspectCommentsUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectCommentsUpdate.open Session("ClientCnnString")
	
	SQLProspectCommentsUpdate = "UPDATE PR_Prospects Set Comments = '" & txtProspectEditComments & "' WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier
	
	Set rsProspectCommentsUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectCommentsUpdate.CursorLocation = 3 
	Set rsProspectCommentsUpdate = cnnProspectCommentsUpdate.Execute(SQLProspectCommentsUpdate)
	
	Description = "The comments for prospect " & ProspectName  & " were changed to <strong><em>" & txtProspectEditComments & "</em></strong> from <strong><em>" & txtProspectCurrentComments & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " prospect comments changed",GetTerm("Prospecting") & " prospect comments changed","Major",0,Description
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	
	set rsProspectCommentsUpdate = Nothing
	cnnProspectCommentsUpdate.Close
	set cnnProspectCommentsUpdate = Nothing
		
End If

Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)
%>