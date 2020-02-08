<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM FS_Parts where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnnparts = Server.CreateObject("ADODB.Connection")
cnnparts.open (Session("ClientCnnString"))
Set rsparts = Server.CreateObject("ADODB.Recordset")
rsparts.CursorLocation = 3 
Set rsparts = cnnparts.Execute(SQL)
	
If not rsparts.EOF Then
	Orig_PartNumber = rsparts("PartNumber")
	Orig_PartDescription = rsparts("PartDescription")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

PartNumber = Request.Form("txtPartNumber")
PartDescription = Request.Form("txtPartDescription")
DisplayOrder = Request.Form("txtPartDisplayOrder")
SearchKeyword = Request.Form("txtSearchKeywords")


SQL = "UPDATE FS_Parts SET "
SQL = SQL &  "PartNumber = '" & PartNumber & "' "
SQL = SQL &  ", PartDescription = '" & PartDescription & "' "
SQL = SQL &  ", DisplayOrder = '" & DisplayOrder & "' "
SQL = SQL &  ", SearchKeywords = '" & SearchKeyword & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

'Response.Write("<br>" & SQL & "<br>")

Set rsparts = cnnparts.Execute(SQL)
set rsparts = Nothing


Description = ""
If Orig_PartNumber  <> PartNumber  Then
	Description = Description & "Service module part number changed from " & Orig_PartNumber  & " to " & PartNumber  
End If

CreateAuditLogEntry "Service module part number edited","Service module part number edited","Minor",0,Description

Response.Redirect("main.asp")

%>















