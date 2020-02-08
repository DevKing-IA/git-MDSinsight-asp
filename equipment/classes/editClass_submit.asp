<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_Classes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Class = rs("Class")
	Orig_RecordSource = rs("RecordSource")
	Orig_InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

EquipmentClass = Request.Form("txtClass")
EquipmentClass = Replace(EquipmentClass, "'", "''")

SQL = "UPDATE EQ_Classes SET "
SQL = SQL &  "Class = '" & EquipmentClass & "', "
SQL = SQL &  "InsightAssetTagPrefix = '" & InsightAssetTagPrefix & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Class <> EquipmentClass Then
	Description = Description & GetTerm("Equipment") & " class changed from " & Orig_Class & " to " & EquipmentClass 
End If
If Orig_InsightAssetTagPrefix <> InsightAssetTagPrefix Then
	Description = Description & GetTerm("Equipment") & " class, " & EquipmentClass & ", changed the Insight Asset Tag Prefix from " & Orig_InsightAssetTagPrefix & " to " & InsightAssetTagPrefix
End If
If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " class Record Source changed from " & Orig_RecordSource & " to " & RecordSource
End If

CreateAuditLogEntry GetTerm("Equipment") & " Class Edited",GetTerm("Equipment") & " Class Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















