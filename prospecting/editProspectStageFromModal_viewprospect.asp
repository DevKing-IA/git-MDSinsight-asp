<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="editProspectStageFromModalSave.asp"-->
<%

If selProspectNextStageNumber = 0 OR selProspectNextStageNumber = 1 Then
	Response.Redirect ("main.asp")
Else
	Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)	
End If

%>