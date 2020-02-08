<!--#include file="../../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<%
ContactTitleNoToReplace = Request.Form("txtContactTitleNoToReplace")
ContactTitleNoReplaceWith = Request.Form("seldeleteContactTitleFromModal")

If ContactTitleNoToReplace <> "" AND ContactTitleNoReplaceWith <> "" Then

	'********************************************************************************************
	'FIRST UPDATE ALL THE TITLES FOR ALL THE PROSPECTS IN PR_PROSPECTS
	'We need to loop through all the records so we can make entries in the PR_Activty table
	'********************************************************************************************
	
	SQLDelete = "Select * From PR_ProspectContacts WHERE ContactTitleNumber = " & ContactTitleNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then
		Do
			Activity = "The contact title for " & rsDelete("FirstName") & " " & rsDelete("LastName") & " was changed from ''" & GetContactTitleByNum(ContactTitleNoToReplace) & "'' to ''" & GetContactTitleByNum(ContactTitleNoReplaceWith) & "'' to allow for the deletion of ''" & GetContactTitleByNum(ContactTitleNoToReplace) & "''"
			Record_PR_Activity rsDelete("ProspectIntRecID"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	
	rsDelete.Close
	
	'Now replace all prospect records with the new contact title number
	
	SQLDelete = "UPDATE PR_ProspectContacts Set ContactTitleNumber = " & ContactTitleNoReplaceWith & " WHERE ContactTitleNumber = " & ContactTitleNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	'********************************************************************************************
	'SECOND, UPDATE ALL THE TITLES FOR ALL THE CUSTOMERS IN AR_CUSTOMERCONTACTS
	'We need to loop through all the records so we can make entries in the PR_Activty table
	'********************************************************************************************
	
	
	SQLDelete = "Select * From AR_CustomerContacts WHERE ContactTitleNumber = " & ContactTitleNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then
		Do
			CustomerIntRecID = rsDelete("CustomerIntRecID")
			CustomerID = GetCustNumByCustIntRecID(CustomerIntRecID)
			Activity = "The contact title for " & rsDelete("FirstName") & " " & rsDelete("LastName") & " customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID & ", was changed from ''" & GetContactTitleByNum(ContactTitleNoToReplace) & "'' to ''" & GetContactTitleByNum(ContactTitleNoReplaceWith) & "'' to allow for the deletion of ''" & GetContactTitleByNum(ContactTitleNoToReplace) & "''"
			CreateAuditLogEntry GetTerm("Customer") & "contact title deleted",GetTerm("Customer") & "  contact title deleted","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new contact title number
	
	SQLDelete = "UPDATE AR_CustomerContacts Set ContactTitleNumber = " & ContactTitleNoReplaceWith & " WHERE ContactTitleNumber = " & ContactTitleNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	

	'********************************************************************************************
	'ONCE BOTH TABLES ARE UPDATED, WE CAN PERFORM THE DELETION
	'********************************************************************************************
	
	SQLDelete = "Delete FROM PR_ContactTitles WHERE InternalRecordIdentifier = "& ContactTitleNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	ContactTitle = GetContactTitleByNum(ContactTitleNoToReplace) ' For audit trail below
	
	Description = "The " & GetTerm("Prospecting") & " and " & GetTerm("Customer") & " contact title named " & ContactTitle & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " and " & GetTerm("Customer") & " contact title deleted",GetTerm("Prospecting") & " and " & GetTerm("Customer") & "  contact title deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>