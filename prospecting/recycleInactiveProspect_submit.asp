<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
txtOrigProspectOwner = Request.Form("txtOrigProspectOwner")
selNewProspectOwner = Request.Form("selNewProspectOwner")
txtOwner = selNewProspectOwner

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	
ProspectIntRecID = txtInternalRecordIdentifier

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
maildomain = Replace(UCASE(maildomain),"WWW.","")

selProspectCurrentActivity = GetCurrentProspectActivityByProspectNumber(txtInternalRecordIdentifier)
selProspectNextActivity = Request.Form("selProspectNextActivity")
txtProspectEditNextActivityNotes = Request.Form("txtNextActivityNotes")
txtProspectEditNextActivityDate = Request.Form("txtNextActivityDueDate")

txtMeetingLocation = Replace(Request.Form("txtMeetingLocation"),"'","''")
selAppointmentDuration = Request.Form("selAppointmentDuration")
selMeetingDuration = Request.Form("selMeetingDuration")

ProspectNewActivity = GetActivityByNum(selProspectNextActivity)
ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(selProspectNextActivity)

txtCurrentStageNo = Request.Form("txtCurrentStageNo")
txtStageNotes = Request.Form("txtStageNotes")
radStageSelected = Request.Form("radStage")


chkDoNotEmailNewOwner = Request.Form("chkDoNotEmailNewOwner")

If (chkDoNotEmailNewOwner <> "" AND chkDoNotEmailNewOwner = "on") Then 
	chkDoNotEmailNewOwner = 1 
	sendEmailFlag = 0
Else 
	chkDoNotEmailNewOwner = 0
	sendEmailFlag = 1
End If

If (cint(Session("UserNo")) <> cint(selNewProspectOwner)) Then
	
	If sendEmailFlag <> True Then
	
		UserNoForCalendarUpdate = selNewProspectOwner
		
	Else ' Send email option is turned on
		
		UserNoForCalendarUpdate = 0
			
	End If	
		
Else
	
	UserNoForCalendarUpdate = Session("UserNO")

End If




'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************

'UPDATE NEXT ACTIVITY FIELDS


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************



If selProspectNextActivity <> "" AND txtInternalRecordIdentifier <> "" Then
	
		
	'Insert new activity
	
	Set cnnProspectNextActivityInsert = Server.CreateObject("ADODB.Connection")
	cnnProspectNextActivityInsert.open Session("ClientCnnString")
		
	If ProspectApptOrMeeting <> "" Then
	
		If ProspectApptOrMeeting ="Appointment" Then
		
			Duration = cint(selAppointmentDuration)
										
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityAppointmentDuration, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & Session("UserNo") & ",1,0," & Duration & ",'" & txtProspectEditNextActivityNotes & "') "
						
		ElseIf ProspectApptOrMeeting ="Meeting" Then
		
			Duration = cint(selMeetingDuration)
									
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityMeetingDuration, ActivityMeetingLocation, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & Session("UserNo") & ",0,1," & Duration & ",'" & txtMeetingLocation & "','" & txtProspectEditNextActivityNotes & "') "
								
		Else
						
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & Session("UserNo") & ",0,0,'" & txtProspectEditNextActivityNotes & "') "		
		
		End If	
	Else
							
		SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, Notes) "
		SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & Session("UserNo") & ",0,0,'" & txtProspectEditNextActivityNotes & "') "		

	End If
	
	'Response.write(SQLProspectNextActivityInsert)

	Set rsProspectNextActivityInsert = Server.CreateObject("ADODB.Recordset")	
	rsProspectNextActivityInsert.CursorLocation = 3 
	Set rsProspectNextActivityInsert = cnnProspectNextActivityInsert.Execute(SQLProspectNextActivityInsert)	
	
	set rsProspectNextActivityInsert = Nothing
	cnnProspectNextActivityInsert.Close
	set cnnProspectNextActivityInsert = Nothing
	
							
	Description = "The next activity " & selProspectCurrentActivity & " for prospect " & ProspectName  & " was changed to " & GetActivityByNum(selProspectNextActivity)  & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with a due date of " & txtProspectEditNextActivityDate 
	CreateAuditLogEntry GetTerm("Prospecting") & " next activity changed",GetTerm("Prospecting") & " next activity changed","Major",0,Description

	Description = "The next activity was set to <strong><em>" & GetActivityByNum(selProspectNextActivity) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with a due date of " & txtProspectEditNextActivityDate  
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
		


End If


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************

'UPDATE PROSPECT STAGE


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************



If radStageSelected = "0" Then

	'******************************************************************
	'Prospect has been set to Unqualified, so we must retreive reasons
	'******************************************************************
	selUnqualifyingReasons = Request.Form("selUnqualifyingReasons")
	selProspectNextStageNumber = 0 
	
ElseIf radStageSelected = "radStageLost" Then

	'**************************************************************
	'Prospect has been set to Lost, so we must retreive reasons
	'**************************************************************
	selLostReasons = Request.Form("selLostReasons")	
	selProspectNextStageNumber = 1
	
ElseIf radStageSelected <> "radStageWon" AND radStageSelected <> "radStageLost" Then

	'**********************************************************************************
	'Prospect has been set to a primary or secondary stage, and is not
	'Unqualified or Lost, so we need to get the stage number
	'**********************************************************************************
	selProspectNextStageNumber = radStageSelected
	
End If


If radStageSelected = "radStageWon" Then
	
	'Update stage to Won - SQL value for stage is 2

	Set cnnProspectStageUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectStageUpdate.open Session("ClientCnnString")
	
	SQLProspectStageUpdate = "INSERT INTO PR_ProspectStages (ProspectRecID, StageRecID, Notes, StageChangedByUserNo) VALUES (" & txtInternalRecordIdentifier & ",2,'" & txtStageNotes & "'," & Session("UserNo") & ") "

	'Response.write(SQLProspectStageUpdate)
	
	Set rsProspectStageUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectStageUpdate.CursorLocation = 3 
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)	

	Description = "The stage " & GetStageByNum(txtCurrentStageNo) & " for prospect " & ProspectName  & " was changed to WON by " & GetUserDisplayNameByUserNo(Session("UserNo"))
	CreateAuditLogEntry GetTerm("Prospecting") & " stage changed",GetTerm("Prospecting") & " stage changed","Major",0,Description

	Description = "The stage was set to WON by " & GetUserDisplayNameByUserNo(Session("UserNo"))
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")

	SQLProspectStageUpdate = "SELECT MAX (InternalRecordIdentifier) AS Expr1 FROM PR_ProspectStages WHERE ProspectRecID = " & txtInternalRecordIdentifier
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)

	If Not rsProspectStageUpdate.EOF Then StageRecordHolder = rsProspectStageUpdate("Expr1")

	set rsProspectStageUpdate = Nothing
	cnnProspectStageUpdate.Close
	set cnnProspectStageUpdate = Nothing
	
	
	'Update Prospect To Customer Status

	Set cnnProspectUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdate.open Session("ClientCnnString")
	
	SQLProspectUpdate = "UPDATE PR_Prospects SET Pool = 'Won', ConvertedToCustomer = 0 WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier

	Set rsProspectUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdate.CursorLocation = 3 
	Set rsProspectUpdate = cnnProspectUpdate.Execute(SQLProspectUpdate)

	set rsProspectUpdate = Nothing
	cnnProspectUpdate.Close
	set cnnProspectUpdate = Nothing
	
	
End If


If radStageSelected <> "" AND txtInternalRecordIdentifier <> "" AND radStageSelected <> "radStageWon" Then
	
	'Update stage

	Set cnnProspectStageUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectStageUpdate.open Session("ClientCnnString")
	
	SQLProspectStageUpdate = "INSERT INTO PR_ProspectStages (ProspectRecID, StageRecID, Notes, StageChangedByUserNo) VALUES (" & txtInternalRecordIdentifier & "," & selProspectNextStageNumber & ",'" & txtStageNotes & "'," & Session("UserNo") & ") "

	Response.write(SQLProspectStageUpdate)
	
	Set rsProspectStageUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectStageUpdate.CursorLocation = 3 
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)
	
	'***************************************************************************************
	'If prospect was marked as Lost or Unqualified, reason codes must be provided
	'***************************************************************************************
	
	
	If selProspectNextStageNumber = 0 Then

		Description = "The stage " & GetStageByNum(txtCurrentStageNo) & " for prospect " & ProspectName  & " was changed to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selUnqualifyingReasons)
		CreateAuditLogEntry GetTerm("Prospecting") & " stage changed",GetTerm("Prospecting") & " stage changed","Major",0,Description
	
		Description = "The stage was set to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selUnqualifyingReasons)
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")

	ElseIf selProspectNextStageNumber = 1 Then

		Description = "The stage " & GetStageByNum(txtCurrentStageNo) & " for prospect " & ProspectName  & " was changed to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selLostReasons)
		CreateAuditLogEntry GetTerm("Prospecting") & " stage changed",GetTerm("Prospecting") & " stage changed","Major",0,Description
	
		Description = "The stage was set to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selLostReasons)
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	Else
		Description = "The stage " & GetStageByNum(txtCurrentStageNo) & " for prospect " & ProspectName  & " was changed to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
		CreateAuditLogEntry GetTerm("Prospecting") & " stage changed",GetTerm("Prospecting") & " stage changed","Major",0,Description
	
		Description = "The stage was set to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	End If	

	SQLProspectStageUpdate = "SELECT MAX (InternalRecordIdentifier) AS Expr1 FROM PR_ProspectStages WHERE ProspectRecID = " & txtInternalRecordIdentifier
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)

	If Not rsProspectStageUpdate.EOF Then StageRecordHolder = rsProspectStageUpdate("Expr1")

	set rsProspectStageUpdate = Nothing
	cnnProspectStageUpdate.Close
	set cnnProspectStageUpdate = Nothing
	
End If



'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************

'DETERMINE IF EMAIL NEEDS TO BE SENT/CALENDAR ENTRIES NEED TO BE MADE
'THEN UPDATE PROSPECT TO LIVE POOL


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'******************************************************************************************************************************************

'Call function to determine email/Outlook actions

response.write("txtInternalRecordIdentifier: " & txtInternalRecordIdentifier & "<br><br>")
response.write("selNewProspectOwner: " & selNewProspectOwner & "<br><br>")
response.write("sendEmailFlag: " & sendEmailFlag & "<br><br>")

dummy = SetOwner_MakeOutlookEntry_SendEmail(txtInternalRecordIdentifier,selNewProspectOwner,sendEmailFlag,"R")




'Update Prospect To Live Pool

Set cnnProspectUpdateLive = Server.CreateObject("ADODB.Connection")
cnnProspectUpdateLive.open Session("ClientCnnString")

SQLProspectUpdateLive = "UPDATE PR_Prospects SET Pool = 'Live' WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier

Set rsProspectUpdateLive = Server.CreateObject("ADODB.Recordset")
rsProspectUpdateLive.CursorLocation = 3 
Set rsProspectUpdateLive = cnnProspectUpdateLive.Execute(SQLProspectUpdateLive)

set rsProspectUpdateLive = Nothing
cnnProspectUpdateLive.Close
set cnnProspectUpdateLive = Nothing

Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " successfully recycled the prospect " & ProspectName & "."	
CreateAuditLogEntry GetTerm("Prospecting") & " prospect recycled",GetTerm("Prospecting") & " prospect recycled","Major",0,Description
Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")



'Insert Log Note


Set cnnProspectLogNote = Server.CreateObject("ADODB.Connection")
cnnProspectLogNote.open Session("ClientCnnString")


SQLProspectLogNote = "INSERT INTO PR_ProspectNotes (ProspectIntRecID, DateAndTime, EnteredByUserNo, Note, Sticky) "
SQLProspectLogNote = SQLProspectLogNote & "VALUES (" & txtInternalRecordIdentifier & ", getdate(), " & Session("Userno") & ", "
SQLProspectLogNote = SQLProspectLogNote & "'Prospect Moved From " & GetTerm("Recycle Pool") & " back to main prospect pool.', 0)"


Set rsProspectLogNote = Server.CreateObject("ADODB.Recordset")
rsProspectLogNote.CursorLocation = 3 
Set rsProspectLogNote = cnnProspectLogNote.Execute(SQLProspectLogNote)

set rsProspectLogNote = Nothing
cnnProspectLogNote.Close
set cnnProspectLogNote = Nothing




Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)
%>