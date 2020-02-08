<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_EmployeeRangeTable where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Range = rs("Range")
	Orig_ProjectedGPSpend = rs("ProjectedGPSpend")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

Range = Request.Form("txtEmployeeRange")
ProjectedGPSpend = Request.Form("txtProjectedGPSpend")

SQL = "UPDATE PR_EmployeeRangeTable SET "
SQL = SQL &  "Range = '" & Range & "', ProjectedGPSpend = " & ProjectedGPSpend & " "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""

If Orig_Range <> Range Then
	Description = Description & GetTerm("Prospecting") & " employee range changed from " & Orig_Range & " to " & Range
End If
CreateAuditLogEntry GetTerm("Prospecting") & " employee range edited",GetTerm("Prospecting") & " employee range edited","Minor",0,Description

If Orig_ProjectedGPSpend <> ProjectedGPSpend Then
	Description = Description & GetTerm("Prospecting") & " employee projected GP Spend changed from " & Orig_ProjectedGPSpend & " to " & ProjectedGPSpend 
End If
CreateAuditLogEntry GetTerm("Prospecting") & " employee range edited",GetTerm("Prospecting") & " employee range projected GP Spend edited","Minor",0,Description


Response.Redirect("main.asp")

%>















