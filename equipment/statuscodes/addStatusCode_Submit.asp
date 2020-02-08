<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

statusDesc = Request.Form("txtStatusCode")
statusDesc = Replace(statusDesc, "'", "''")

If Request.Form("chkAvailableForPlacement") = "on" then statusAvailableForPlacement = 1 Else statusAvailableForPlacement = 0
If Request.Form("chkGeneratesRentalRevenue") = "on" then statusGeneratesRentalRevenue = 1 Else statusGeneratesRentalRevenue = 0
  
statusBackendSystemCode = Request.Form("txtBackendSystemCode")


SQL = "INSERT INTO EQ_StatusCodes (statusDesc,statusAvailableForPlacement,statusGeneratesRentalRevenue,statusBackendSystemCode,RecordSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & statusDesc & "'," & statusAvailableForPlacement & "," & statusGeneratesRentalRevenue& ",'" & statusBackendSystemCode & "','Insight')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " status code: " & statusDesc
CreateAuditLogEntry GetTerm("Equipment") & " Status Code Added",GetTerm("Equipment") & " Status Code Added","Minor",0,Description

Response.Redirect("main.asp")

%>















