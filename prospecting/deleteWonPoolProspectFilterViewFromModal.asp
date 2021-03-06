﻿<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->

<%

reportNameToDelete = Request("viewNameToDelete")

If UCASE(reportNameToDelete) = "CURRENT" OR UCASE(reportNameToDelete) = "DEFAULT" OR UCASE(reportNameToDelete) = "ALL PROSPECTS" Then
	reportNameToDelete = ""
End If

If reportNameToDelete <> "" Then
	
	dummy = MUV_WRITE("CRMVIEWSTATEWONPOOL","Default")
	
	reportNameToDeleteSQL = Replace(reportNameToDelete,"'","''")
		
	Set cnnReportSettings = Server.CreateObject("ADODB.Connection")
	cnnReportSettings.open (Session("ClientCnnString"))
	
	Set rsReportSettings = Server.CreateObject("ADODB.Recordset")
	rsReportSettings.CursorLocation = 3 

	SQLDeleteFilter = "DELETE FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Won' AND UserReportName = '" & reportNameToDeleteSQL & "'"
	Set rsReportSettings = cnnReportSettings.Execute(SQLDeleteFilter)

	cnnReportSettings.Close
	Set rsReportSettings = Nothing
	Set cnnReportSettings = Nothing

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " deleted the " & GetTerm("Prospecting") & "  " & GetTerm("New Customer Pool") & "filter view named, " & reportNameToDeleteSQL
	CreateAuditLogEntry GetTerm("Prospecting") & " filter view deleted",GetTerm("Prospecting") & " filter view deleted","Minor",0,Description
	
End If

Response.Redirect ("mainWonPool.asp")

%>