<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_Reasons where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Reason = rs("Reason")
	Orig_ReasonType = rs("ReasonType")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

Reason = Request.Form("txtReason")
ReasonType = Request.Form("selReasonType")

SQL = "UPDATE PR_Reasons SET "
SQL = SQL &  "Reason = '" & Reason & "',ReasonType = '" & ReasonType & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""


If Orig_Reason <> Reason AND Orig_ReasonType <> ReasonType Then

	Description = Description & GetTerm("Prospecting") & " reason changed from " & Orig_Reason & " to " & Reason & " and reason type changed from " & Orig_ReasonType & " to " & ReasonType
Else

	If Orig_Reason  <> Reason  Then
		Description = Description & GetTerm("Prospecting") & " reason changed from " & Orig_Reason & " to " & Reason
	End If
	
	If Orig_ReasonType <> ReasonType  Then
		Description = Description & GetTerm("Prospecting") & " reason type from " & Orig_ReasonType & " to " & ReasonType
	End If
	
End If


CreateAuditLogEntry GetTerm("Prospecting") & " reason edited",GetTerm("Prospecting") & " reason edited","Minor",0,Description

Response.Redirect("main.asp")

%>















