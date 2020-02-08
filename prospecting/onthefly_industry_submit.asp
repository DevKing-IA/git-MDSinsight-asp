<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

Industry = Request.Form("txtIndustry")
InternalRecordIdentifier = Request.Form("txtpid") 

'check if fields are not empty
If Industry<>"" Then
	Industry = Hacker_Filter2(Industry)
End If

SQL = "INSERT INTO PR_Industries (Industry)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Industry & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " industry, on the fly: " & Industry 
CreateAuditLogEntry GetTerm("Prospecting") & " industry added",GetTerm("Prospecting") & " industry added","Minor",0,Description

'Response.Redirect("viewProspectDetail.asp?i=" & InternalRecordIdentifier)
%>
