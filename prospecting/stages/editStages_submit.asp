<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_Stages where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_StagesDescription = rs("Stage")
	Orig_StageType = rs("StageType")
	Orig_SortOrder = rs("SortOrder")
	Orig_ProbabilityPercent = rs("ProbabilityPercent")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

StageDescription = Request.Form("txtDescription")
StageType = Request.Form("selStageType")
StageSortOrder = Request.Form("selStageSortOrder")
ProbabilityPercent = Request.Form("selProbabilityPercent")

SQL = "UPDATE PR_Stages SET "
SQL = SQL &  "Stage = '" & StageDescription & "', "
SQL = SQL &  "StageType = '" & StageType & "', "
SQL = SQL &  "SortOrder = " & StageSortOrder & ", "
SQL = SQL &  "ProbabilityPercent = " & ProbabilityPercent & " "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""

If Orig_StagesDescription  <> StagesDescription  Then
	Description = Description & GetTerm("Prospecting") & " stage descrtiption changed from " & Orig_StagesDescription & " to " & StagesDescription 
End If
If Orig_StageType <> StageType Then
	Description = Description & GetTerm("Prospecting") & " stage type changed from " & Orig_StageType & " to " & StageType 
End If
If cdbl(Orig_ProbabilityPercent) <> cdbl(ProbabilityPercent) Then
	Description = Description & GetTerm("Prospecting") & " stage probability percent changed from " & Orig_ProbabilityPercent & "% to " & ProbabilityPercent & "%"
End If
If cdbl(Orig_SortOrder) <> cdbl(StageSortOrder) Then
	Description = Description & GetTerm("Prospecting") & " stage sort order changed from " & Orig_ProbabilityPercent & " to " & StageSortOrder
End If


CreateAuditLogEntry GetTerm("Prospecting") & " stage edited",GetTerm("Prospecting") & " stage edited","Minor",0,Description

Response.Redirect("main.asp")

%>















