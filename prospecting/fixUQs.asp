<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%
Response.write("<BR><BR><BR>")
'Build temp table with all prospects in live pool
Set cnnFixUQs = Server.CreateObject("ADODB.Connection")
cnnFixUQs.open Session("ClientCnnString")
Set rsFixUQs = Server.CreateObject("ADODB.Recordset")
rsFixUQs.CursorLocation = 3 
Set rsFixUQs2 = Server.CreateObject("ADODB.Recordset")
rsFixUQs.CursorLocation = 3 


SQLFixUQs = "DROP TABLE _FixUQs"

On error resume next
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)	
On error goto 0

'Put them all in
SQLFixUQs = "SELECT * INTO _FixUQs FROM PR_Prospects"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)	

'Delete any not in the Live Pool
SQLFixUQs = "DELETE FROM _FixUQs WHERE POOL <> 'Live'"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)

'Reset Stage no on all records
SQLFixUQs = "UPDATE _FixUQs SET StageNo = 9999"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)

'Set Stages to the proper stage number
SQLFixUQs = "SELECT * FROM _FixUQs"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)

Do While NOT rsFixUqs.Eof

	ProspRecID = rsFixUQs("InternalRecordIdentifier")

	SQLFixUQs = "UPDATE _FixUQs Set StageNo = " & GetProspectCurrentStageByProspectNumber(ProspRecID) & " WHERE InternalRecordIdentifier = " & ProspRecID
	Set rsFixUQs2 = cnnFixUQs.Execute(SQLFixUQs)
	
	rsFixUQs.Movenext
Loop

'Get rid of anything where the stage is not 0
SQLFixUQs = "DELETE FROM _FixUQs WHERE StageNo <> 0"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)


'Now do all the main work
SQLFixUQs = "SELECT * FROM _FixUQs"
Response.write(SQLFixUQs & "<BR>")
Set rsFixUQs = cnnFixUQs.Execute(SQLFixUQs)

Do While NOT rsFixUqs.Eof


	selProspectNextStageNumber = 0 
	selUnqualifyingReasons = 0
	
	txtInternalRecordIdentifier =rsFixUQs("InternalRecordIdentifier")
	txtCurrentStageNo = 0
	txtProspectEditStageNotes = "To correct UQs which remained live following salesforce import"

	ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	

	Response.Write("<br> ProspectName: " & ProspectName )






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
	
	
	Description = "The stage " & GetStageByNum(txtCurrentStageNo) & " for prospect " & ProspectName  & " was changed to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selUnqualifyingReasons)
	CreateAuditLogEntry GetTerm("Prospecting") & " stage changed",GetTerm("Prospecting") & " stage changed","Major",0,Description
	
	Description = "The stage was set to " & GetStageByNum(selProspectNextStageNumber) & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with the reason: " & GetReasonByNum(selUnqualifyingReasons)
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")


	SQLProspectStageUpdate = "SELECT MAX (InternalRecordIdentifier) AS Expr1 FROM PR_ProspectStages WHERE ProspectRecID = " & txtInternalRecordIdentifier
	Set rsProspectStageUpdate = cnnProspectStageUpdate.Execute(SQLProspectStageUpdate)

	If Not rsProspectStageUpdate.EOF Then StageRecordHolder = rsProspectStageUpdate("Expr1")

	set rsProspectStageUpdate = Nothing
	cnnProspectStageUpdate.Close
	set cnnProspectStageUpdate = Nothing
	

Response.Write("<br> PROCESS <br>")


'**************************************************************************************************
'If the prospect was marked as Unqualified or Lost, update the Reasons table with the reason code
'**************************************************************************************************



	'Update Reason
	
	Set cnnProspectReasonUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectReasonUpdate.open Session("ClientCnnString")
	
	SQLProspectReasonUpdate = "INSERT INTO PR_ProspectReasons(ProspectRecID, StageRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecId) VALUES (" & txtInternalRecordIdentifier & "," & selProspectNextStageNumber & "," & selUnqualifyingReasons & "," & Session("UserNo") & ", " & StageRecordHolder & ") "


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

	Response.Write("<br> ProspectName: FINISHED")

	rsFixUQs.Movenext
Loop


Response.Write("<br>END END END END END END END END END END END END END END END END END END END END END END END END END END END END END END <br>")
Response.End

%>