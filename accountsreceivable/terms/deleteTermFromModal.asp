<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
TermToBeDeletedIntRecID = Request.Form("txtTermToBeDeletedIntRecID")
TermToReplaceWithIntRecID = Request.Form("selDeleteTermFromModal")

'TermToBeDeleted = GetClassCodeByIntRecID(TermToBeDeletedIntRecID)
TermToBeDeleted = TermToBeDeletedIntRecID
'TermToBeReplacedWith = GetClassCodeByIntRecID(TermToReplaceWithIntRecID)
TermToBeReplacedWith = TermToReplaceWithIntRecID

TermToBeDeletedDescription = GetCustTermDescByIntRecID(TermToBeDeletedIntRecID)


If TermToBeDeleted <> "" AND TermToBeReplacedWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "SELECT CustNum, Name FROM AR_Customer WHERE CustType = '" & TermToBeDeleted & "'"
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	If not rsDelete.Eof Then
		Do
			CustAcctNo = rsDelete("CustNum")
			CustName = rsDelete("Name")
			Description = "The Cust Type for customer " & CustAcctNo & " (" & CustName & ") was changed from ''" & TermToBeDeleted & "'' to ''" & TermToBeReplacedWith & "'' to allow for the deletion of ''" & TermToBeDeleted & "'' by " & GetUserDisplayNameByUserNo(Session("UserNo"))
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer term updated",GetTerm("Accounts Receivable") & " customer term updated","Major",0,Description
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new type code
	
	SQLDelete = "UPDATE AR_Customer SET TermsIntRecID = '" & TermToBeReplacedWith & "' WHERE TermsIntRecID = '" & TermToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM AR_Terms WHERE InternalRecordIdentifier = '" & TermToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	Description = "The " & GetTerm("Accounts Receivable") & " term named " & TermToBeDeletedDescription & " (" & TermToBeDeleted & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Accounts Receivable") & " term deleted",GetTerm("Accounts Receivable") & " term deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>