<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

ProblemDescription = Request.Form("txtProblemDescription")
ShowOnWebsite = Request.Form("selShowOnWeb")

SQL = "INSERT INTO FS_ProblemCodes (ProblemDescription,ShowOnWebsite)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ProblemDescription & "'," & ShowOnWebsite  & ")"

Response.Write("<br>" & SQL & "<br>")

Set cnnProblemCodes = Server.CreateObject("ADODB.Connection")
cnnProblemCodes.open (Session("ClientCnnString"))

Set rsProblemCodes = Server.CreateObject("ADODB.Recordset")
rsProblemCodes.CursorLocation = 3 

Set rsProblemCodes = cnnProblemCodes.Execute(SQL)
set rsProblemCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module Problem code, on the fly: " & ProblemDescription 
CreateAuditLogEntry "Service module" & " Problem code added","Service module Problem code added","Minor",0,Description

%>
