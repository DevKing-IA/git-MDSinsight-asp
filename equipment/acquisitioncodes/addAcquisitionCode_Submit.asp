<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

acquisitionCode = Request.Form("txtAcquisitionCode")
acquisitionCode = Replace(acquisitionCode, "'", "''")

acquisitionCodeDesc = Request.Form("txtAcquisitionCodeDesc")
acquisitionCodeDesc = Replace(acquisitionCodeDesc, "'", "''")


SQL = "INSERT INTO EQ_AcquisitionCodes (acquisitionCode, acquisitionDesc)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & acquisitionCode & "','" & acquisitionCodeDesc & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Acquisition Code: " & acquisitionCode & " (" & acquisitionCodeDesc & ")"
CreateAuditLogEntry GetTerm("Equipment") & " Acquisition Code Added",GetTerm("Equipment") & " Acquisition Code Added","Minor",0,Description

Response.Redirect("main.asp")

%>















