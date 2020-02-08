<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->

<%
SQL = "DELETE FROM Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = 'Current'"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

dummy = MUV_WRITE("CRMVIEWSTATERECPOOL","Default")

Response.Redirect ("mainRecyclePool.asp")
%>

 
