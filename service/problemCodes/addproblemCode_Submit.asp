<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

ProblemDescription = Request.Form("txtProblemDescription")
ShowOnWebsite = Request.Form("selShowOnWeb")

SQL = "INSERT INTO FS_ProblemCodes (ProblemDescription,ShowOnWebsite)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ProblemDescription & "'," & ShowOnWebsite  & ")"

Response.Write("<br>" & SQL & "<br>")

Set cnnproblemCodes = Server.CreateObject("ADODB.Connection")
cnnproblemCodes.open (Session("ClientCnnString"))

Set rsproblemCodes = Server.CreateObject("ADODB.Recordset")
rsproblemCodes.CursorLocation = 3 

Set rsproblemCodes = cnnproblemCodes.Execute(SQL)
set rsproblemCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module problem code: " & ProblemDescription 
CreateAuditLogEntry "Service module" & " problem code added","Service module problem code added","Minor",0,Description

Response.Redirect("main.asp")

%>















