<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

movementCode = Request.Form("txtMovementCode")
movementCode = Replace(movementCode, "'", "''")

movementCodeDesc = Request.Form("txtMovementCodeDesc")
movementCodeDesc = Replace(movementCodeDesc, "'", "''")


SQL = "INSERT INTO EQ_MovementCodes (movementCode, movementDesc)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & movementCode & "','" & movementCodeDesc & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Movement Code: " & movementCode & " (" & movementCodeDesc & ")"
CreateAuditLogEntry GetTerm("Equipment") & " Movement Code Added",GetTerm("Equipment") & " Movement Code Added","Minor",0,Description

Response.Redirect("main.asp")

%>















