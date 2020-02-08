<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

InternalRecordIdentifier = Request.Form("txtpid") 
StageDescription = Request.Form("txtStage")
ProbabilityPercent = Request.Form("selProbabilityPercent")

SQL = "INSERT INTO PR_Stages (Stage,ProbabilityPercent)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & StageDescription & "'," & ProbabilityPercent & ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " stage, on the fly: " & StageDescription 
CreateAuditLogEntry GetTerm("Prospecting") & " Stage Added",GetTerm("Prospecting") & " Stage Added","Minor",0,Description

Response.Redirect("viewProspectDetail.asp?i=" & InternalRecordIdentifier)
%>
