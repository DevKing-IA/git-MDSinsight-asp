<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
ClassCodeToBeDeletedIntRecID = Request.Form("txtClassCodeToBeDeletedIntRecID")
ClassCodeToReplaceWithIntRecID = Request.Form("selDeleteClassCodeFromModal")

ClassCodeToBeDeleted = GetClassCodeByIntRecID(ClassCodeToBeDeletedIntRecID)
ClassCodeToBeReplacedWith = GetClassCodeByIntRecID(ClassCodeToReplaceWithIntRecID)
ClassCodeToBeDeletedDescription = GetCustClassDescByIntRecID(ClassCodeToBeDeletedIntRecID)


If ClassCodeToBeDeleted <> "" AND ClassCodeToBeReplacedWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "SELECT CustNum, Name FROM AR_Customer WHERE ClassCode = '" & ClassCodeToBeDeleted & "'"
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	If not rsDelete.Eof Then
		Do
			CustAcctNo = rsDelete("CustNum")
			CustName = rsDelete("Name")
			Description = "The class code for customer " & CustAcctNo & " (" & CustName & ") was changed from ''" & ClassCodeToBeDeleted & "'' to ''" & ClassCodeToBeReplacedWith & "'' to allow for the deletion of ''" & ClassCodeToBeDeleted & "'' by " & GetUserDisplayNameByUserNo(Session("UserNo"))
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer class code updated",GetTerm("Accounts Receivable") & " customer class code updated","Major",0,Description
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new class code
	
	SQLDelete = "UPDATE AR_Customer SET ClassCode = '" & ClassCodeToBeReplacedWith & "' WHERE ClassCode = '" & ClassCodeToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM AR_CustomerClass WHERE ClassCode = '" & ClassCodeToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	Description = "The " & GetTerm("Accounts Receivable") & " customer class named " & ClassCodeToBeDeletedDescription & " (" & ClassCodeToBeDeleted & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer class deleted",GetTerm("Accounts Receivable") & " customer class deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>