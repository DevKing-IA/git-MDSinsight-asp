<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

TypeDescription = Request.Form("txtCustDescription")

IvsComment1 = Request.Form("txtIvsComment1")
IvsComment2 = Request.Form("txtIvsComment2")
IvsComment3 = Request.Form("txtIvsComment3")
IvsComment4 = Request.Form("txtIvsComment4")
IvsComment5 = Request.Form("txtIvsComment5")
HoldDays = Request.Form("txtHoldDays")
HoldAmt = Request.Form("txtHoldAmt")
WholesaleFlag = Request.Form("txtWholesaleFlag")
MemoMessagingFlag = Request.Form("txtMemoMessagingFlag")

SQL = "INSERT INTO AR_CustomerType (TypeDescription, IvsComment1, IvsComment2, IvsComment3, IvsComment4, IvsComment5, HoldDays, HoldAmt, WholesaleFlag, MemoMessagingFlag)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & TypeDescription & "',"
SQL = SQL & "'"  & IvsComment1 & "',"
SQL = SQL & "'"  & IvsComment2 & "',"
SQL = SQL & "'"  & IvsComment3 & "',"
SQL = SQL & "'"  & IvsComment4 & "',"
SQL = SQL & "'"  & IvsComment5 & "',"
SQL = SQL & "'"  & HoldDays & "',"
SQL = SQL & "'"  & HoldAmt & "',"
SQL = SQL & "'"  & WholesaleFlag & "',"
SQL = SQL & "'"  & MemoMessagingFlag & "')"



'Response.Write("<br>" & SQL & "<br>")

Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))

Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the accounts receivable module customer type: " & TypeDescription 
CreateAuditLogEntry "Accounts Receivable module" & " customer type added","Accounts Receivable module customer type added","Minor",0,Description

Response.Redirect("main.asp")

%>















