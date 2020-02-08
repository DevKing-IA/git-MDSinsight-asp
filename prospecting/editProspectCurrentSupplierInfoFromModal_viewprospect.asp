<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
txtProspectCurrentCurrentOffering = Request.Form("txtProspectCurrentCurrentOffering")
txtProspectEditCurrentOffering = Request.Form("txtProspectEditCurrentOffering")

If txtProspectEditCurrentOffering = "" Then txtProspectEditCurrentOffering = "CURRENT OFFERING REMOVED"
If txtProspectCurrentCurrentOffering = "" Then txtProspectCurrentCurrentOffering = "NONE ENTERED"

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	

If txtInternalRecordIdentifier <> "" Then
	
	'Update prospect CurrentOffering

	Set cnnProspectCurrentOfferingUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectCurrentOfferingUpdate.open Session("ClientCnnString")
	
	SQLProspectCurrentOfferingUpdate = "UPDATE PR_Prospects Set CurrentOffering = '" & txtProspectEditCurrentOffering & "' WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier
	
	Set rsProspectCurrentOfferingUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectCurrentOfferingUpdate.CursorLocation = 3 
	Set rsProspectCurrentOfferingUpdate = cnnProspectCurrentOfferingUpdate.Execute(SQLProspectCurrentOfferingUpdate)
	
	Description = "The Current Offering for prospect " & ProspectName  & " were changed to <strong><em>" & txtProspectEditCurrentOffering & "</em></strong> from <strong><em>" & txtProspectCurrentCurrentOffering & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " prospect current offering changed",GetTerm("Prospecting") & " prospect current offering changed","Major",0,Description
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	
	set rsProspectCurrentOfferingUpdate = Nothing
	cnnProspectCurrentOfferingUpdate.Close
	set cnnProspectCurrentOfferingUpdate = Nothing
		
End If

Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)
%>