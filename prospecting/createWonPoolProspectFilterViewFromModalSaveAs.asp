<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

viewNameInputField = Request("viewNameInputField")
viewNameSelectBox = Request("viewNameSelectBox")

If viewNameInputField = "" Then
	reportNameToSave = viewNameSelectBox
Else
	reportNameToSave = viewNameInputField
End If

If reportNameToSave <> "" Then

	
	dummy = MUV_WRITE("CRMVIEWSTATEWONPOOL",reportNameToSave)
	
	reportNameToSaveSQL = Replace(reportNameToSave,"'","''")
		
	Set cnnReportSettings = Server.CreateObject("ADODB.Connection")
	cnnReportSettings.open (Session("ClientCnnString"))
	
	Set rsReportSettings = Server.CreateObject("ADODB.Recordset")
	rsReportSettings.CursorLocation = 3 

	SQLDeleteFilter = "DELETE FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Won' AND UserReportName = '" & reportNameToSaveSQL & "'"
	Set rsReportSettings = cnnReportSettings.Execute(SQLDeleteFilter)
	
	SQLReportSettings = "UPDATE Settings_Reports SET UserReportName = '" & reportNameToSaveSQL & "' WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Won' AND UserReportName='Current'"
	Set rsReportSettings = cnnReportSettings.Execute(SQLReportSettings)

	cnnReportSettings.Close
	Set rsReportSettings = Nothing
	Set cnnReportSettings = Nothing

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " created the " & GetTerm("Prospecting") & "  " & GetTerm("New Customer Pool") & "filter view named, " & reportNameToSave
	CreateAuditLogEntry GetTerm("Prospecting") & " filter view added",GetTerm("Prospecting") & " filter view added","Minor",0,Description
	

End If

%>