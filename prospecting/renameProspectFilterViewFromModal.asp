<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

newViewName = Request("newViewName")
newViewNameSaveSQL = Replace(newViewName,"'","''")

originalViewName = Request("originalViewName")
originalViewNameSaveSQL = Replace(originalViewName,"'","''")

If newViewName <> "" AND originalViewName <> "" Then
	
	dummy = MUV_WRITE("CRMVIEWSTATE",newViewName)	

	Set cnnReportSettings = Server.CreateObject("ADODB.Connection")
	cnnReportSettings.open (Session("ClientCnnString"))
	Set rsReportSettings = Server.CreateObject("ADODB.Recordset")
	rsReportSettings.CursorLocation = 3 
	
	SQLReportSettings = "UPDATE Settings_Reports SET UserReportName = '" & newViewNameSaveSQL & "' WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & originalViewNameSaveSQL & "'"	
	Set rsReportSettings = cnnReportSettings.Execute(SQLReportSettings)

	cnnReportSettings.Close
	Set rsReportSettings = Nothing
	Set cnnReportSettings = Nothing


	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " renamed the " & GetTerm("Prospecting") & " filter view named, " & originalViewNameSaveSQL & ", to " & newViewNameSaveSQL & "."
	CreateAuditLogEntry GetTerm("Prospecting") & " filter view renamed",GetTerm("Prospecting") & " filter view renamed","Minor",0,Description
	
End If

%>