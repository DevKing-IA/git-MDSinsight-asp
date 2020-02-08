<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

ResolutionDescription = Request.Form("txtResolutionDescription")

SQL = "INSERT INTO FS_ResolutionCodes (ResolutionDescription)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ResolutionDescription & "')"

Response.Write("<br>" & SQL & "<br>")

Set cnnresolutionCodes = Server.CreateObject("ADODB.Connection")
cnnresolutionCodes.open (Session("ClientCnnString"))

Set rsresolutionCodes = Server.CreateObject("ADODB.Recordset")
rsresolutionCodes.CursorLocation = 3 

Set rsresolutionCodes = cnnresolutionCodes.Execute(SQL)
set rsresolutionCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module resolution code: " & ResolutionDescription 
CreateAuditLogEntry "Service module" & " resolution code added","Service module resolution code added","Minor",0,Description

Response.Redirect("main.asp")

%>















