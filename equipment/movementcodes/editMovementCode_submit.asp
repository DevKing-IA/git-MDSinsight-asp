<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

movementCode = Request.Form("txtMovementCode")
movementCode = Replace(movementCode, "'", "''")

movementCodeDesc = Request.Form("txtMovementCodeDesc")
movementCodeDesc = Replace(movementCodeDesc, "'", "''")


'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_MovementCodes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_movementCode = rs("movementCode")
	Orig_movementCodeDesc = rs("movementDesc")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_MovementCodes SET "
SQL = SQL &  "movementCode = '" & movementCode & "', "
SQL = SQL &  "movementDesc = '" & movementCodeDesc & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_movementCode <> movementCode Then
	Description = Description & GetTerm("Equipment") & " Movement Code changed from " & Orig_movementCode & " to " & movementCode 
End If

If Orig_movementCodeDesc <> movementCodeDesc Then
	Description = Description & GetTerm("Equipment") & " Movement Code description changed from " & Orig_movementCodeDesc & " to " & movementCodeDesc
End If

CreateAuditLogEntry GetTerm("Equipment") & " Condition Edited",GetTerm("Equipment") & " Movement Code Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















