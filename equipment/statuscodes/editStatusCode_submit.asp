<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

statusDesc = Request.Form("txtStatusCode")
statusDesc = Replace(statusDesc, "'", "''")

If Request.Form("chkAvailableForPlacement") = "on" then statusAvailableForPlacement = 1 Else statusAvailableForPlacement = 0
If Request.Form("chkGeneratesRentalRevenue") = "on" then statusGeneratesRentalRevenue = 1 Else statusGeneratesRentalRevenue = 0

statusBackendSystemCode = Request.Form("txtBackendSystemCode")


'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_StatusCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_statusDesc = rs("statusDesc")
	Orig_statusBackendSystemCode = rs("statusBackendSystemCode")
	Orig_statusAvailableForPlacement = rs("statusAvailableForPlacement")
	Orig_statusGeneratesRentalRevenue = rs("statusGeneratesRentalRevenue")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_StatusCodes SET "
SQL = SQL &  "statusDesc = '" & statusDesc & "', "
SQL = SQL &  "statusAvailableForPlacement = " & statusAvailableForPlacement & ", "
SQL = SQL &  "statusGeneratesRentalRevenue = " & statusGeneratesRentalRevenue & ", "
SQL = SQL &  "statusBackendSystemCode = '" & statusBackendSystemCode & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_statusDesc <> statusDesc Then
	Description = Description & GetTerm("Equipment") & " status code description changed from " & Orig_statusDesc & " to " & statusDesc 
End If

If cInt(statusAvailableForPlacement) = 1 then statusAvailableForPlacementONOFFMsg = "On" Else statusAvailableForPlacementONOFFMsg = "Off"
If cInt(Orig_statusAvailableForPlacement) = 1 then OrigStatusAvailableForPlacementONOFFMsg = "On" Else OrigStatusAvailableForPlacementONOFFMsg = "Off"

IF cInt(Orig_statusAvailableForPlacement) <> cInt(statusAvailableForPlacement) Then
	Description = Description & GetTerm("Equipment") & " status code description changed from " & OrigStatusAvailableForPlacementONOFFMsg & " to " & statusAvailableForPlacementONOFFMsg
End If

If cInt(statusGeneratesRevenue) = 1 then statusGeneratesRevenueONOFFMsg = "On" Else statusGeneratesRevenueONOFFMsg = "Off"
If cInt(Orig_statusGeneratesRevenue) = 1 then OrigstatusGeneratesRevenueONOFFMsg = "On" Else OrigstatusGeneratesRevenueONOFFMsg = "Off"

IF cInt(Orig_statusGeneratesRevenue) <> cInt(statusGeneratesRevenue) Then
	Description = Description & GetTerm("Equipment") & " status code generates rental revenue changed from " & OrigstatusGeneratesRevenueONOFFMsg & " to " & statusGeneratesRevenueONOFFMsg
End If


If Orig_statusBackendSystemCode <> statusBackendSystemCode Then
	Description = Description & GetTerm("Equipment") & " status code backend system code changed from " & Orig_statusBackendSystemCode & " to " & statusBackendSystemCode
End If

If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " status code Record Source changed from " & Orig_RecordSource & " to " & RecordSource
End If

CreateAuditLogEntry GetTerm("Equipment") & " Condition Edited",GetTerm("Equipment") & " Status Code Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















