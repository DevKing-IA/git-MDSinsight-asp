<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_Condition where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Condition = rs("Condition")
	Orig_ConditionDescription = rs("Description")
	Orig_RecordSource = rs("RecordSource")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

Condition = Request.Form("txtCondition")
Condition = Replace(Condition, "'", "''")

ConditionDescription = Request.Form("txtConditionDescription")
ConditionDescription = Replace(ConditionDescription, "'", "''")

SQL = "UPDATE EQ_Condition SET "
SQL = SQL &  "Condition = '" & Condition & "', "
SQL = SQL &  "Description = '" & ConditionDescription & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Condition  <> Condition  Then
	Description = Description & GetTerm("Equipment") & " condition changed from " & Orig_Condition & " to " & Condition
End If
If Orig_ConditionDescription <> ConditionDescription Then
	Description = Description & GetTerm("Equipment") & " condition description changed from " & Orig_ConditionDescription & " to " & ConditionDescription
End If
If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " condition Record Source changed from " & Orig_RecordSource & " to " & RecordSource
End If

CreateAuditLogEntry GetTerm("Equipment") & " Condition Edited",GetTerm("Equipment") & " Condition Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















