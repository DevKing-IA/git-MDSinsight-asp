<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

Range1 = Request.Form("txtEmployeeRange1")
Range2 = Request.Form("txtEmployeeRange2")
ProjectedGPSpend = Request.Form("txtProjectedGPSpend")

Range = Range1 & "-" & Range2

SQL = "INSERT INTO PR_EmployeeRangeTable (Range, ProjectedGPSpend)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Range & "'," & ProjectedGPSpend & ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " employee range: " & Range & " with projected GP Spend of " & ProjectedGPSpend
CreateAuditLogEntry GetTerm("Prospecting") & " employee range added",GetTerm("Prospecting") & " employee range added","Minor",0,Description

Response.Redirect("main.asp")

%>















