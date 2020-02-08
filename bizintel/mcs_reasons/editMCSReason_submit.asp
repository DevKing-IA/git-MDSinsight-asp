<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"-->
<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM BI_MCSReasons WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Reason = rs("Reason")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

Reason = Request.Form("txtMCSReason")
Reason = Replace(Reason, "'", "''")

SQL = "UPDATE BI_MCSReasons SET Reason = '" & Reason & "' WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""

If Orig_Reason <> Reason  Then
	Description = Description & GetTerm("Business Intelligence") & " MCS reason changed from " & Orig_Reason & " to " & Reason
End If
	
CreateAuditLogEntry GetTerm("Business Intelligence") & " MCS reason edited",GetTerm("Business Intelligence") & " MCS reason edited","Minor",0,Description

Response.Redirect("main.asp")

%>















