<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%

ProspectIDArray = Split(Request.Form("prospectsArray"),",")
	
Set rsDelete = Server.CreateObject("ADODB.Recordset")
rsDelete.CursorLocation = 3 
Set cnnDelete = Server.CreateObject("ADODB.Connection")
cnnDelete.open (Session("ClientCnnString"))


For i = 0 to uBound(ProspectIDArray)

	ProspectIDNumber = cInt(ProspectIDArray(i))
	ProspectName = GetProspectNameByNumber(ProspectIDNumber)
	
	SQLDelete = "DELETE FROM PR_ProspectStages where ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectReasons WHERE ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectEmailLog WHERE ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectDocuments WHERE ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectContacts WHERE ProspectIntRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectCompetitors WHERE ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectNotes WHERE ProspectIntRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_ProspectActivities WHERE ProspectRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	SQLDelete = "DELETE FROM PR_Audit WHERE ProspectIntRecID = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	SQLDelete = "DELETE FROM PR_Prospects WHERE InternalRecordIdentifier = " & ProspectIDNumber
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " deleted the prospect " & ProspectName
	CreateAuditLogEntry "Prospect deleted","Prospect deleted","Major",0,Description

	
Next

Set rsDelete = Nothing
cnnDelete.Close
Set cnnDelete= Nothing

Response.Redirect ("main.asp")

%>