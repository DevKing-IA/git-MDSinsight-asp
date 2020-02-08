<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
CustChainToBeDeletedIntRecID = Request.Form("txtChainCodeToBeDeletedIntRecID")
'CustChainToReplaceWithIntRecID = Request.Form("selDeleteRefCodeFromModal")

'CustChainToBeDeleted = GetClassCodeByIntRecID(CustChainToBeDeletedIntRecID)
CustChainToBeDeleted = CustChainToBeDeletedIntRecID
'CustChainToBeReplacedWith = GetClassCodeByIntRecID(CustChainToReplaceWithIntRecID)
CustChainToBeReplacedWith = 0

CustChainToBeDeletedDescription = GetCustChainByIntRecID(CustChainToBeDeletedIntRecID)


If CustChainToBeDeleted <> "" AND CustChainToBeReplacedWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "SELECT CustNum, Name FROM AR_Customer WHERE ReferalCode = '" & CustChainToBeDeleted & "'"
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	If not rsDelete.Eof Then
		Do
			CustAcctNo = rsDelete("CustNum")
			CustName = rsDelete("Name")
			Description = "The Chain for customer " & CustAcctNo & " (" & CustName & ") was changed from ''" & CustChainToBeDeleted & "'' to ''" & CustChainToBeReplacedWith & "'' to allow for the deletion of ''" & CustChainToBeDeleted & "'' by " & GetUserDisplayNameByUserNo(Session("UserNo"))
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer Chain updated",GetTerm("Accounts Receivable") & " customer chain updated","Major",0,Description
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new referal code
	
	SQLDelete = "UPDATE AR_Customer SET ChainNum = '" & CustChainToBeReplacedWith & "' WHERE ChainNum = '" & CustChainToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM AR_Chain WHERE InternalRecordIdentifier = '" & CustChainToBeDeleted & "'"
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	'Response.write(SQLDelete & "<br>")
	
	Description = "The " & GetTerm("Accounts Receivable") & " customer chain named " & CustChainToBeDeletedDescription & " (" & CustChainToBeDeleted & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer chain deleted",GetTerm("Accounts Receivable") & " customer chain deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>