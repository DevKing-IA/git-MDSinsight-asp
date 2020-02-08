<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_CustomerType where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))
Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 
Set rscust = cnncust.Execute(SQL)
	
If not rscust.EOF Then	
	Orig_TypeDescription = rscust("TypeDescription")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

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


SQL = "UPDATE AR_CustomerType SET "
SQL = SQL &  "TypeDescription = '" & TypeDescription & "', "
SQL = SQL &  "IvsComment1 = '" & IvsComment1 & "', "
SQL = SQL &  "IvsComment2 = '" & IvsComment2 & "', "
SQL = SQL &  "IvsComment3 = '" & IvsComment3 & "', "
SQL = SQL &  "IvsComment4 = '" & IvsComment4 & "', "
SQL = SQL &  "IvsComment5 = '" & IvsComment5 & "', "
SQL = SQL &  "HoldDays = '" & HoldDays & "', "
SQL = SQL &  "HoldAmt = '" & HoldAmt & "', "
SQL = SQL &  "WholesaleFlag = '" & WholesaleFlag & "', "
SQL = SQL &  "MemoMessagingFlag = '" & MemoMessagingFlag & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

'Response.Write("<br>" & SQL & "<br>")

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing


Description = ""
If Orig_TypeDescription  <> TypeDescription  Then
	Description = Description & "Accounts Receivable module customer type changed from " & Orig_TypeDescription  & " to " & TypeDescription  
End If

CreateAuditLogEntry "Accounts Receivable module customer type edited","Accounts Receivable module customer type edited","Minor",0,Description

Response.Redirect("main.asp")

%>















