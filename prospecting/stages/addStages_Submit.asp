<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

StageDescription = Request.Form("txtDescription")
StageType = Request.Form("selStageType")
StageSortOrder = Request.Form("selStageSortOrder")
ProbabilityPercent = Request.Form("selProbabilityPercent")

SQL = "INSERT INTO PR_Stages (Stage,StageType,ProbabilityPercent,SortOrder)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & StageDescription & "','"  & StageType & "'," & ProbabilityPercent & "," & StageSortOrder & ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " stage: " & StageDescription & " of type " & StageType & " with a sort order of " & StageSortOrder 
CreateAuditLogEntry GetTerm("Prospecting") & " Stage Added",GetTerm("Prospecting") & " Stage Added","Minor",0,Description

Response.Redirect("main.asp")

%>















