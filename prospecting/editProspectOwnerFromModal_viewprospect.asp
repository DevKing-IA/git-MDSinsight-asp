<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
txtCurrentProspectOwner = Request.Form("txtCurrentProspectOwner")
selProspectEditOwner = Request.Form("selProspectEditOwner")
txtOwner = selProspectEditOwner

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	
ProspectIntRecID = txtInternalRecordIdentifier

chkDoNotEmailNewOwner = Request.Form("chkDoNotEmailNewOwner")

If (chkDoNotEmailNewOwner <> "" AND chkDoNotEmailNewOwner = "on") Then 
	chkDoNotEmailNewOwner = 1 
	sendEmailFlag = 0
Else 
	chkDoNotEmailNewOwner = 0
	sendEmailFlag = 1
End If


'''dummy = SetOwner_MakeOutlookEntry_SendEmail(passedProspectID,passedNewOwnerUserNo,passedSendEmailFlag,passedPageSource)
''''passedPageSource: R = Recycle, E = Edit Prospect, A = Add Prospect ,O = Owner Request Email Accepted???


dummy = SetOwner_MakeOutlookEntry_SendEmail(txtInternalRecordIdentifier,selProspectEditOwner,sendEmailFlag,"E")


Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)
%>