<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

SymptomDescription = Request.Form("txtSymptomDescription")
ShowOnWebsite = Request.Form("selShowOnWeb")

SQL = "INSERT INTO FS_SymptomCodes (SymptomDescription,ShowOnWebsite)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & SymptomDescription & "'," & ShowOnWebsite  & ")"

Response.Write("<br>" & SQL & "<br>")

Set cnnSymptomCodes = Server.CreateObject("ADODB.Connection")
cnnSymptomCodes.open (Session("ClientCnnString"))

Set rsSymptomCodes = Server.CreateObject("ADODB.Recordset")
rsSymptomCodes.CursorLocation = 3 

Set rsSymptomCodes = cnnSymptomCodes.Execute(SQL)
set rsSymptomCodes = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module Symptom code, on the fly: " & SymptomDescription 
CreateAuditLogEntry "Service module" & " Symptom code added","Service module Symptom code added","Minor",0,Description

%>
