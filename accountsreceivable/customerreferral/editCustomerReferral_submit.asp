<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_CustomerReferral where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))
Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 
Set rscust = cnncust.Execute(SQL)
	
If not rscust.EOF Then	
	Orig_ReferralName = rscust("ReferralName")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

ReferralName = Request.Form("txtReferralName")

RefDescription = Request.Form("txtCustDescription")
RefDescription2 = Request.Form("txtCustDescription2")


SQL = "UPDATE AR_CustomerReferral SET "
SQL = SQL &  "ReferralName = '" & ReferralName & "', "
SQL = SQL &  "Description = '" & RefDescription & "', "
SQL = SQL &  "Description2 = '" & RefDescription2 & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

'Response.Write("<br>" & SQL & "<br>")

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing


Description = ""
If Orig_ReferralName  <> ReferralName  Then
	Description = Description & "Accounts Receivable module customer referral changed from " & Orig_ReferralName  & " to " & ReferralName  
End If

CreateAuditLogEntry "Accounts Receivable module customer referral edited","Accounts Receivable module customer referral edited","Minor",0,Description

Response.Redirect("main.asp")

%>















