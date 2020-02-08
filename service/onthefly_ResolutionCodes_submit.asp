<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

ResolutionDescription = Request.Form("txtResolutionDescription")

SQL = "INSERT INTO FS_ResolutionCodes (ResolutionDescription)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ResolutionDescription & "')"

Response.Write("<br>" & SQL & "<br>")

Set cnnResolutionCodes = Server.CreateObject("ADODB.Connection")
cnnResolutionCodes.open (Session("ClientCnnString"))

Set rsResolutionCodes = Server.CreateObject("ADODB.Recordset")
rsResolutionCodes.CursorLocation = 3 

Set rsResolutionCodes = cnnResolutionCodes.Execute(SQL)
set rsResolutionCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module Resolution code, on the fly : " & ResolutionDescription 
CreateAuditLogEntry "Service module" & " Resolution code added","Service module Resolution code added","Minor",0,Description

%>
