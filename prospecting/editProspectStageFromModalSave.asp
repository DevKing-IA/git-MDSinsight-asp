
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
txtCurrentStageNo = Request.Form("txtCurrentStageNo")
txtProspectEditStageNotes = Request.Form("txtProspectEditStageNotes")
radStageSelected = Request.Form("radStage")


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

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	


If radStageSelected = "radStageWon" Then
	
	'Update stage to Won - SQL value for stage is 2

	Set cnnProspectStageUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectStageUpdate.open Session("ClientCnnString")
	
	SQLProspectStageUpdate = "INSERT INTO PR_ProspectStages (ProspectRecID, StageRecID, Notes, StageChangedByUserNo) VALUES (" & txtInternalRecordIdentifier & ",2,'" & txtProspectEditStageNotes & "'," & Session("UserNo") & ") "

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
	
	SQLProspectStageUpdate = "INSERT INTO PR_ProspectStages (ProspectRecID, StageRecID, Notes, StageChangedByUserNo) VALUES (" & txtInternalRecordIdentifier & "," & selProspectNextStageNumber & ",'" & txtProspectEditStageNotes & "'," & Session("UserNo") & ") "

	Response.write(SQLProspectStageUpdate)
	
	Set rsProspectStageUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectStageUpdate.CursorLocation = 3 
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)
	
	'***************************************************************************************
	'If prospect was marked as Lost or Unqualified, reason codes must be provided
	'***************************************************************************************
	
	'Response.write("selUnqualifyingReasons : " & selUnqualifyingReasons & "<br><br>")
	'Response.write("selLostReasons : " & selLostReasons & "<br><br>")
	'Response.write("selProspectNextStageNumber : " & selProspectNextStageNumber & "<br><br>")
	
	
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

'**************************************************************************************************
'If the prospect was marked as Unqualified or Lost, update the Reasons table with the reason code
'**************************************************************************************************
If (selProspectNextStageNumber = 0 OR selProspectNextStageNumber = 1) AND radStageSelected <> "radStageWon" Then

	'Update Reason
	
	Set cnnProspectReasonUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectReasonUpdate.open Session("ClientCnnString")
	
	If selProspectNextStageNumber = 0 Then
		SQLProspectReasonUpdate = "INSERT INTO PR_ProspectReasons(ProspectRecID, StageRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecId) VALUES (" & txtInternalRecordIdentifier & "," & selProspectNextStageNumber & "," & selUnqualifyingReasons & "," & Session("UserNo") & ", " & StageRecordHolder & ") "
	ElseIf selProspectNextStageNumber = 1 Then
		SQLProspectReasonUpdate = "INSERT INTO PR_ProspectReasons(ProspectRecID, StageRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecID) VALUES (" & txtInternalRecordIdentifier & "," & selProspectNextStageNumber & "," & selLostReasons & "," & Session("UserNo") & ", " & StageRecordHolder & ") "
	End If
	

	Set rsProspectReasonUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectReasonUpdate.CursorLocation = 3 
	Set rsProspectReasonUpdate = cnnProspectReasonUpdate.Execute(SQLProspectReasonUpdate)

	set rsProspectReasonUpdate = Nothing
	cnnProspectReasonUpdate.Close
	set cnnProspectReasonUpdate = Nothing
	
	
	'Update Prospect Next Activity Status
	

	Set cnnProspectActivityUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectActivityUpdate.open Session("ClientCnnString")
	
	SQLProspectNextActivityUpdate = "UPDATE PR_ProspectActivities Set Status = 'Cancelled',StatusDateTime = GetDate(), "
	SQLProspectNextActivityUpdate = SQLProspectNextActivityUpdate & " Notes = 'Moved To " & GetTerm("Dead Pool") & "', StatusChangedByUserNo = " & Session("UserNo")
	SQLProspectNextActivityUpdate = SQLProspectNextActivityUpdate & " WHERE ProspectRecID = " & txtInternalRecordIdentifier & " AND Status IS NULL "

	Set rsProspectActivityUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectActivityUpdate.CursorLocation = 3 
	Set rsProspectActivityUpdate = cnnProspectActivityUpdate.Execute(SQLProspectNextActivityUpdate)

	set rsProspectActivityUpdate = Nothing
	cnnProspectActivityUpdate.Close
	set cnnProspectActivityUpdate = Nothing
	

	
	'Update Prospect To Dead Pool

	Set cnnProspectUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdate.open Session("ClientCnnString")
	
	SQLProspectUpdate = "UPDATE PR_Prospects SET Pool = 'Dead' WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier

	Set rsProspectUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdate.CursorLocation = 3 
	Set rsProspectUpdate = cnnProspectUpdate.Execute(SQLProspectUpdate)

	set rsProspectUpdate = Nothing
	cnnProspectUpdate.Close
	set cnnProspectUpdate = Nothing


End If 

%>