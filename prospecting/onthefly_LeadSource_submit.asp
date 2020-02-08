<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

LeadSource = Request.Form("txtLeadSource")
Industry = Request.Form("txtIndustry")
InternalRecordIdentifier = Request.Form("txtpid") 

If LeadSource<>"" Then
	LeadSource = Hacker_Filter2(LeadSource)
End If

SQL = "INSERT INTO PR_LeadSources (LeadSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & LeadSource & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " lead source, on the fly: " & LeadSource 
CreateAuditLogEntry GetTerm("Prospecting") & " lead source added",GetTerm("Prospecting") & " lead source added","Minor",0,Description

'Response.Redirect("viewProspectDetail.asp?i=" & InternalRecordIdentifier)
%>

