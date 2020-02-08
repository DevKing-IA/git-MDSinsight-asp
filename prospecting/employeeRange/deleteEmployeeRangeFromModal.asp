<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
EmployeeRangeNoToReplace = Request.Form("txtEmployeeRangeNoToReplace")
EmployeeRangeNoReplaceWith = Request.Form("seldeleteEmployeeRangeFromModal")

If EmployeeRangeNoToReplace <> "" AND EmployeeRangeNoReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "Select InternalRecordIdentifier From PR_Prospects WHERE EmployeeRangeNumber = " & EmployeeRangeNoToReplace 
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The reason this employee range was changed from ''" & GetEmployeeRangeByNum(EmployeeRangeNoToReplace) & "'' to ''" & GetEmployeeRangeByNum(EmployeeRangeNoReplaceWith) & "'' to allow for the deletion of ''" & GetEmployeeRangeByNum(EmployeeRangeNoToReplace) & "''"
	ReasonDescription = GetEmployeeRangeByNum(EmployeeRangeNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Reason number
	
	SQLDelete = "UPDATE PR_Prospects Set EmployeeRangeNumber = " & EmployeeRangeNoReplaceWith & " WHERE  EmployeeRangeNumber = " & EmployeeRangeNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_EmployeeRangeTable WHERE InternalRecordIdentifier = "& EmployeeRangeNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " employee range " & ReasonDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " employee range deleted",GetTerm("Prospecting") & " employee range deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>