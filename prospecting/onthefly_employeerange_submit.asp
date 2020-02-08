<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

InternalRecordIdentifier = Request.Form("txtpid") 

Range1 = Request.Form("txtEmployeeRange1")
Range2 = Request.Form("txtEmployeeRange2")

Range = Range1 & "-" & Range2

SQL = "INSERT INTO PR_EmployeeRangeTable (Range)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Range & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " employee range, on the fly: " & Range
CreateAuditLogEntry GetTerm("Prospecting") & " employee range added",GetTerm("Prospecting") & " employee range added","Minor",0,Description

'Response.Redirect("viewProspectDetail.asp?i=" & InternalRecordIdentifier)
%>
