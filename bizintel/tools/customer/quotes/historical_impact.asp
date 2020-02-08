<%
MonthToEvaluate = selNumMonthsHistoricalImpact 
If MonthToEvaluate = "" Then MonthToEvaluate = 6
Set cnnHistImpact  = Server.CreateObject("ADODB.Connection")
cnnHistImpact.open (Session("ClientCnnString"))
Set rsHistImpact = Server.CreateObject("ADODB.Recordset")
rsHistImpact.CursorLocation = 3 

On Error Resume Next ' In case the table isn't there
SQLHistImpact = "DROP TABLE zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno"))
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)
On Error Goto 0

SQLHistImpact = "CREATE TABLE zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno"))
SQLHistImpact = SQLHistImpact & "("
SQLHistImpact = SQLHistImpact &	"                [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
SQLHistImpact = SQLHistImpact & "                [prodSKU] [varchar](255) NULL, "
SQLHistImpact = SQLHistImpact & "                [UM] [varchar](255) NULL, "
SQLHistImpact = SQLHistImpact & "                [Qty] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [HistoricalCost] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [HistoricalPrice] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [quotedCost] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [quotedPrice] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [newPrice] [decimal](18,2) NULL, "
SQLHistImpact = SQLHistImpact & "                [quoted] [int] NULL "
SQLHistImpact = SQLHistImpact & ")"
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)


' Insert all the products and units Historically sold to this cust in the last XX months where the price is not 0
SQLHistImpact = "INSERT INTO zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " (prodSKU, UM, Qty, HistoricalCost, HistoricalPrice, quotedCost, quotedPrice, quoted) "
SQLHistImpact = SQLHistImpact & "SELECT partNum, prodSalesUnit,  itemQuantity, itemCOst, itemPrice, 0, 0,  1 FROM InvoiceHistoryDetail WHERE "
SQLHistImpact = SQLHistImpact &	"(CustNum = '" & CustID & "') AND (ivsDate > DATEADD(m, - " & MonthToEvaluate & ", GETDATE() - DATEPART(d, GETDATE()) + 1)) "
SQLHistImpact = SQLHistImpact &	"AND itemPrice <> 0"
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

' Now reflag as not quoted any items not currently quoted to this account
SQLHistImpact = "UPDATE zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " "
SQLHistImpact = SQLHistImpact & "SET Quoted = 0 "
SQLHistImpact = SQLHistImpact & "WHERE prodSKU NOT IN "
SQLHistImpact = SQLHistImpact & "(SELECT prodSKU FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ")"
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

'Fill in quoted prices for these items and quoted costs
SQLHistImpact = "UPDATE zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " "
SQLHistImpact = SQLHistImpact & "SET zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".quotedPrice = "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".Price, "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".quotedCost = "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".Cost FROM "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " INNER JOIN zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " ON "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".prodSKU = zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".prodSKU AND "			
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".QuoteType = zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".UM"									
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

'Fill in all the new price fields
SQLHistImpact = "UPDATE zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " "
SQLHistImpact = SQLHistImpact & "SET zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".newPrice = "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".NewPrice FROM "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & " INNER JOIN zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " ON "
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".prodSKU = zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".prodSKU AND "			
SQLHistImpact = SQLHistImpact & "zPRC_AccountQuotedItems_" & trim(Session("Userno")) & ".QuoteType = zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) & ".UM"									
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)


'Calculate Historical figures
TotRevenueHistorical = 0 
TotalCostHistorical = 0
SQLHistImpact = "SELECT SUM (Qty*HistoricalPrice) as TotPrice, SUM (Qty*HistoricalCost) as TotCost FROM zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) 
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

If not rsHistImpact.Eof Then
	TotRevenueHistorical = rsHistImpact("TotPrice")
	TotalCostHistorical = rsHistImpact("TotCost")
End If

'Calculate Quoted figures
TotRevenueCurrent = 0
TotCostCurrent = 0
SQLHistImpact = "SELECT * FROM zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) 
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

If not rsHistImpact.Eof Then

	Do WHile not rsHistImpact.Eof
	
		If rsHistImpact("quoted") = 1 Then
			TotRevenueCurrent = TotRevenueCurrent + (cdbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("QuotedPrice"))) 
		Else 
			TotRevenueCurrent = TotRevenueCurrent + (cdbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("historicalPrice"))) 
		End If
		
		TotCostCurrent = TotCostCurrent + (cDbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("historicalCost")))

		rsHistImpact.MoveNext
	Loop
	
End If


'Calculate projected figures
TotRevenueProjected = 0
TotalCostProjected = 0
SQLHistImpact = "SELECT * FROM zPRC_AccountQuotedItems_Impact_" & trim(Session("Userno")) 
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

If not rsHistImpact.Eof Then

	Do WHile not rsHistImpact.Eof
	
		If rsHistImpact("quoted") = 1 Then
			If Not IsNull(rsHistImpact("newPrice")) Then
				TotRevenueProjected = TotRevenueProjected + (cDbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("newPrice"))) 
			Else
				TotRevenueProjected  = TotRevenueProjected  + (cdbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("QuotedPrice"))) 
			End If
		Else ' Not a quoted item, use original price
			TotRevenueProjected  = TotRevenueProjected  + (cdbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("historicalPrice"))) 
		End If
		
		TotalCostProjected = TotalCostProjected + (cDbl(rsHistImpact("Qty")) * cDbl(rsHistImpact("historicalCost")))

		rsHistImpact.MoveNext
	Loop
	
End If


' Write all the good info we just came up wiht


Response.Write("<table border='0' class='table table-striped table-bordered' cellspacing='0'>")
Response.Write("<tr>")
Response.Write("<td colspan='2'>Based on cost of " & FormatCurrency(TotalCostHistorical,2) & "</td>")
Response.Write("<td>Sales</td>")
Response.Write("<td>GP $</td>")
Response.Write("<td>GP %</td>")
Response.Write("</tr>")

Response.Write("<tr>")
Response.Write("<td colspan='2'>Actual data</td>")
Response.Write("<td>" & FormatCurrency(TotRevenueHistorical,2) & "</td>")
Response.Write("<td>" & FormatCurrency(cdbl(TotRevenueHistorical) - cdbl(TotalCostHistorical),2) & "</td>")
Response.Write("<td>" & Round(((cdbl(TotRevenueHistorical) - cdbl(TotalCostHistorical)) / cdbl(TotRevenueHistorical)) * 100,2) & "%</td>")
Response.Write("</tr>")

Response.Write("<tr>")
Response.Write("<td colspan='2'>At current prices</td>")
Response.Write("<td>" & FormatCurrency(TotRevenueCurrent,2) & "</td>")
Response.Write("<td>" & FormatCurrency(cdbl(TotRevenueCurrent) - cdbl(TotCostCurrent),2) & "</td>")
Response.Write("<td>" & Round(((cdbl(TotRevenueCurrent) - cdbl(TotCostCurrent)) / cdbl(TotRevenueCurrent)) * 100,2) & "%</td>")
Response.Write("</tr>")

Response.Write("<tr>")
Response.Write("<td colspan='2'>At new prices</td>")
Response.Write("<td>" & FormatCurrency(TotRevenueProjected,2) & "</td>")
Response.Write("<td>" & FormatCurrency(cdbl(TotRevenueProjected) - cdbl(TotalCostProjected),2) & "</td>")
Response.Write("<td>" & Round(((cdbl(TotRevenueProjected) - cdbl(TotalCostProjected)) / cdbl(TotRevenueProjected)) * 100,2) & "%</td>")
Response.Write("</tr>")

'Calculate differences
DiffTotRevenueProjected = TotRevenueProjected - TotRevenueCurrent
DiffTotGPDollars = ((cdbl(TotRevenueProjected) - cdbl(TotalCostProjected)))- ((cdbl(TotRevenueCurrent) - cdbl(TotCostCurrent)))
DiffTotGPPercent = ((((cdbl(TotRevenueProjected) - cdbl(TotalCostProjected)) / cdbl(TotRevenueProjected)) * 100)) - ((((cdbl(TotRevenueCurrent) - cdbl(TotCostCurrent)) / cdbl(TotRevenueCurrent)) * 100))

Response.Write("<tr>")
Response.Write("<td colspan='2'><strong>Difference current vs. new</strong></td>")
tdclass= ""
If DiffTotRevenueProjected > 0 Then tdclass = " class='highlight-green' "
If DiffTotRevenueProjected < 0 Then tdclass = " class='highlight-red' "
If DiffTotRevenueProjected = 0 Then tdclass = " class='highlight-orange-ish' "
If tdclass = "" Then
	Response.Write("<td id='DiffTotRevenueProjected'><strong>" & FormatCurrency(DiffTotRevenueProjected,2) & "</strong></td>")
Else
	Response.Write("<td id='DiffTotRevenueProjected' " & tdclass & "><strong>" & FormatCurrency(DiffTotRevenueProjected,2) & "</strong></td>")
End If
tdclass= ""
If DiffTotGPDollars > 0 Then tdclass = " class='highlight-green' "
If DiffTotGPDollars < 0 Then tdclass = " class='highlight-red' "
If DiffTotGPDollars = 0 Then tdclass = " class='highlight-orange-ish' "
If tdclass = "" Then
	Response.Write("<td id='DiffTotGPDollars'><strong>" & FormatCurrency(DiffTotGPDollars ,2) & "</strong></td>")
Else
	Response.Write("<td id='DiffTotGPDollars' " & tdclass & "><strong>" & FormatCurrency(DiffTotGPDollars ,2) & "</strong></td>")
End If
tdclass= ""
If DiffTotGPPercent > 0 Then tdclass = " class='highlight-green' "
If DiffTotGPPercent < 0 Then tdclass = " class='highlight-red' "
If DiffTotGPPercent = 0 Then tdclass = " class='highlight-orange-ish' "
If tdclass = "" Then
	Response.Write("<td id='DiffTotGPPercent'><strong>" & Round(DiffTotGPPercent ,2) & "%</strong></td>")
Else
	Response.Write("<td id='DiffTotGPPercent' " & tdclass & "><strong>" & Round(DiffTotGPPercent ,2) & "%</strong></td>")
End If
Response.Write("</tr>")

Response.Write("</table>")

Set rsHistImpact = Nothing
cnnHistImpact.Close
Set cnnHistImpact = Nothing	   

'*********************************************************************
'Now use this information to build the historical impact summary table
'*********************************************************************
Set cnnHistImpact  = Server.CreateObject("ADODB.Connection")
cnnHistImpact.open (Session("ClientCnnString"))
Set rsHistImpact = Server.CreateObject("ADODB.Recordset")
rsHistImpact.CursorLocation = 3 
Set rsHistImpact2 = Server.CreateObject("ADODB.Recordset")
rsHistImpact2.CursorLocation = 3 
Set rsHistImpact3 = Server.CreateObject("ADODB.Recordset")
rsHistImpact3.CursorLocation = 3 



On Error Resume Next ' In caase the table isn't there
SQLHistImpact = "DROP TABLE zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno"))
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)
On Error Goto 0

SQLHistImpact = "CREATE TABLE zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno"))
SQLHistImpact = SQLHistImpact & "("
SQLHistImpact = SQLHistImpact &	"                [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
SQLHistImpact = SQLHistImpact & "                [prodSKU] [varchar](255) NULL, "
SQLHistImpact = SQLHistImpact & "                [UM] [varchar](255) NULL, "
SQLHistImpact = SQLHistImpact & "                [NumTimesOrdered] [int] NULL, "
SQLHistImpact = SQLHistImpact & "                [MostRecentOrdDate] [datetime] NULL "
SQLHistImpact = SQLHistImpact & ")"
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

' Insert all the products and units and the number of times ordered in the last XX mos WITH THE PRICE, EVEN IF $0
SQLHistImpact = "INSERT INTO zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno")) & " (prodSKU, UM, NumTimesOrdered) "
SQLHistImpact = SQLHistImpact & "SELECT partNum, prodSalesUnit,  COUNT(partnum) AS OrdCount FROM InvoiceHistoryDetail WHERE "
SQLHistImpact = SQLHistImpact &	"(CustNum = '" & CustID & "') AND (ivsDate > DATEADD(m, - " & MonthToEvaluate & ", GETDATE() - DATEPART(d, GETDATE()) + 1)) "
SQLHistImpact = SQLHistImpact &	"GROUP BY partnum, prodSalesUnit"
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

'Update all dates to NULL to get rid of the wacky 1/1/1900 values
SQLHistImpact = "UPDATE zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno")) & " "
SQLHistImpact = SQLHistImpact & "SET MostRecentOrdDate = NULL "
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)


SQLHistImpact = "SELECT prodSku, UM FROM zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno")) 
Set rsHistImpact = cnnHistImpact.Execute(SQLHistImpact)

If not rsHistImpact.Eof Then

	Do WHile not rsHistImpact.Eof
	
		SQLHistImpact2 = "SELECT TOP 1 IvsDate FROM InvoiceHistoryDetail Where CustNum = '" & CustID & "' AND partnum = '" & rsHistImpact("prodSKU") & "' "
		SQLHistImpact2 = SQLHistImpact2 & " AND prodSalesUnit = '" & rsHistImpact("UM") & "' ORDER BY IvsDate DESC"
		Set rsHistImpact2 = cnnHistImpact.Execute(SQLHistImpact2)
		
		If Not rsHistImpact2.EOF Then
		
			SQLHistImpact3 = "UPDATE zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno")) 
			SQLHistImpact3 = SQLHistImpact3 & " SET MostRecentOrdDate = '" & rsHistImpact2("IvsDate") & "' WHERE prodSKU='" & rsHistImpact("prodSKU") & "' AND UM = '" & rsHistImpact("UM") & "'"
			Set rsHistImpact3 = cnnHistImpact.Execute(SQLHistImpact3)
		
		End If
		
		rsHistImpact.MoveNext
	Loop
	
End If

Set rsHistImpact3 = Nothing
Set rsHistImpact2 = Nothing
Set rsHistImpact = Nothing
cnnHistImpact.Close
Set cnnHistImpact = Nothing	  
%>
