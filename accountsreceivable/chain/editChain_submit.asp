<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_Chain where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))
Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 
Set rscust = cnncust.Execute(SQL)
	
If not rscust.EOF Then	
	Orig_Description = rscust("Description")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

ChainDescription = Request.Form("txtChainDescription")
UpdateDiscount = Request.Form("txtUpdateDiscount")
SellOnlyQuoted = Request.Form("txtSellOnlyQuoted")
ChainPrice = Request.Form("txtChainPrice")
PoFlag = Request.Form("txtPoFlag")
PurchaseOrder = Request.Form("txtPurchaseOrder")
ProgramType = Request.Form("txtProgramType")
PrimarySalesman = Request.Form("txtPrimarySalesman")
WebRequiredFields = Request.Form("txtWebRequiredFields")
DefQuoteValidDate = Request.Form("txtDefQuoteValidDate")


SQL = "UPDATE AR_Chain SET "
SQL = SQL &  "Description = '" & ChainDescription & "', "
SQL = SQL &  "updateDiscount = '" & UpdateDiscount & "', "
SQL = SQL &  "SellOnlyQuoted = '" & SellOnlyQuoted & "', "
SQL = SQL &  "chainPrice = '" & ChainPrice & "', "
SQL = SQL &  "poFlag = '" & PoFlag & "', "
SQL = SQL &  "purchaseOrder = '" & PurchaseOrder & "', "
SQL = SQL &  "programType = '" & ProgramType & "', "
SQL = SQL &  "primarySalesman = '" & PrimarySalesman & "', "
SQL = SQL &  "webRequiredFields = '" & WebRequiredFields & "', "
SQL = SQL &  "defQuoteValidDate = '" & DefQuoteValidDate & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

'Response.Write("<br>" & SQL & "<br>")

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing


Description = ""
If Orig_Description  <> ChainDescription  Then
	Description = Description & "Accounts Receivable module chain changed from " & Orig_Description  & " to " & ChainDescription  
End If

CreateAuditLogEntry "Accounts Receivable module chain edited","Accounts Receivable module chain edited","Minor",0,Description

Response.Redirect("main.asp")

%>















