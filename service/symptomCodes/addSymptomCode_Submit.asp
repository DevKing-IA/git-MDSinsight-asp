<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

SymptomDescription = Request.Form("txtSymptomDescription")
ShowOnWebsite = Request.Form("selShowOnWeb")

SQL = "INSERT INTO FS_SymptomCodes (SymptomDescription,ShowOnWebsite)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & SymptomDescription & "'," & ShowOnWebsite  & ")"

Response.Write("<br>" & SQL & "<br>")

Set cnnsymptomCodes = Server.CreateObject("ADODB.Connection")
cnnsymptomCodes.open (Session("ClientCnnString"))

Set rssymptomCodes = Server.CreateObject("ADODB.Recordset")
rssymptomCodes.CursorLocation = 3 

Set rssymptomCodes = cnnsymptomCodes.Execute(SQL)
set rssymptomCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module symptom code: " & SymptomDescription 
CreateAuditLogEntry "Service module" & " symptom code added","Service module symptom code added","Minor",0,Description

Response.Redirect("main.asp")

%>















