<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

ProblemDescription = Request.Form("txtProblemDescription")
ShowOnWebsite = Request.Form("selShowOnWeb")

PartNumber = Request.Form("txtPartNumber")
PartDescription = Request.Form("txtPartDescription")
DisplayOrder = Request.Form("txtPartDisplayOrder")
SearchKeyword = Request.Form("txtSearchKeywords")


If DisplayOrder = "" Then
	DisplayOrder = 0
End If	

SQL = "INSERT INTO FS_Parts (PartNumber,PartDescription,DisplayOrder,SearchKeywords)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & PartNumber & "', '"  & PartDescription & "', " & DisplayOrder  & ", '"  & SearchKeyword & "')"

Response.Write("<br>" & SQL & "<br>")

Set cnnparts = Server.CreateObject("ADODB.Connection")
cnnparts.open (Session("ClientCnnString"))

Set rsparts = Server.CreateObject("ADODB.Recordset")
rsparts.CursorLocation = 3 

Set rsparts = cnnparts.Execute(SQL)
set rsparts = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the service module part number: " & ProblemDescription 
CreateAuditLogEntry "Service module" & " part number added","Service module part number added","Minor",0,Description

Response.Redirect("main.asp")

%>















