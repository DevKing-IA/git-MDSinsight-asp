<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fill in the audit trail

SQL = "SELECT * FROM EQ_Groups where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Group = rs("GroupName")
	Orig_RecordSource = rs("RecordSource")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

EquipmentGroup = Request.Form("txtGroup")
EquipmentGroup = Replace(EquipmentGroup, "'", "''")

SQL = "UPDATE EQ_Groups SET "
SQL = SQL &  "GroupName = '" & EquipmentGroup & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Group  <> EquipmentGroup Then
	Description = Description & GetTerm("Equipment") & " Group changed from " & Orig_Group & " to " & EquipmentGroup 
End If
If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " Group Record Source changed from " & Orig_RecordSource & " to " & RecordSource 
End If


CreateAuditLogEntry GetTerm("Equipment") & " Group Edited",GetTerm("Equipment") & " Group Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















