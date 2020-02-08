<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

acquisitionCode = Request.Form("txtacquisitionCode")
acquisitionCode = Replace(acquisitionCode, "'", "''")

acquisitionCodeDesc = Request.Form("txtacquisitionCodeDesc")
acquisitionCodeDesc = Replace(acquisitionCodeDesc, "'", "''")


'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_AcquisitionCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_acquisitionCode = rs("acquisitionCode")
	Orig_acquisitionCodeDesc = rs("acquisitionDesc")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_AcquisitionCodes SET "
SQL = SQL &  "acquisitionCode = '" & acquisitionCode & "', "
SQL = SQL &  "acquisitionDesc = '" & acquisitionCodeDesc & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_acquisitionCode <> acquisitionCode Then
	Description = Description & GetTerm("Equipment") & " Acquisition Code changed from " & Orig_acquisitionCode & " to " & acquisitionCode 
End If

If Orig_acquisitionCodeDesc <> acquisitionCodeDesc Then
	Description = Description & GetTerm("Equipment") & " Acquisition Code description changed from " & Orig_acquisitionCodeDesc & " to " & acquisitionCodeDesc
End If

CreateAuditLogEntry GetTerm("Equipment") & " Condition Edited",GetTerm("Equipment") & " Acquisition Code Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















