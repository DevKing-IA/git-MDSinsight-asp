<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

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

SQL = "INSERT INTO AR_Chain (Description, updateDiscount, SellOnlyQuoted, chainPrice, poFlag, purchaseOrder, programType, primarySalesman, webRequiredFields, defQuoteValidDate, mtdQty0, mtdAmt0, mtdQty1, mtdAmt1, mtdQty2, mtdAmt2, mtdQty3, mtdAmt3, mtdQty4, mtdAmt4, mtdQty5, mtdAmt5, mtdQty6, mtdAmt6, mtdQty7, mtdAmt7, mtdQty8, mtdAmt8, mtdQty9, mtdAmt9, mtdQty10, mtdAmt10, mtdQty11, mtdAmt11, qcd0, qcd1, qcd2, qcd3, qcd4, qcd5, qcd6, qcd7, qcd8, qcd9, qcd10, qcd11, qcd12, qcd13, qcd14, qcd15, qcd16, qcd17, qcd18, qcd19, qcd20)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ChainDescription & "',"
SQL = SQL & "'"  & UpdateDiscount & "',"
SQL = SQL & "'"  & SellOnlyQuoted & "',"
SQL = SQL & "'"  & ChainPrice & "',"
SQL = SQL & "'"  & PoFlag & "',"
SQL = SQL & "'"  & PurchaseOrder & "',"
SQL = SQL & "'"  & ProgramType & "',"
SQL = SQL & "'"  & PrimarySalesman & "',"
SQL = SQL & "'"  & WebRequiredFields & "',"
SQL = SQL & "'"  & DefQuoteValidDate & "', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)"


'Response.Write("<br>" & SQL & "<br>")

Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))

Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the accounts receivable module chain: " & ChainDescription 
CreateAuditLogEntry "Accounts Receivable module" & " chain added","Accounts Receivable module chain added","Minor",0,Description

Response.Redirect("main.asp")

%>















