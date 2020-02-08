<style>
	.smaller-detail-line{
		font-size: 0.8em;
	}	
</style>
<%

TotalCustsReported = 0

emailbody = "<!DOCTYPE html>" & vbcrlf
emailbody = emailbody & "<!--[if lt IE 7 ]> <html class='no-js ie6 oldie' lang='en'> <![endif]-->" & vbcrlf
emailbody = emailbody & "<!--[if IE 7 ]>    <html class='no-js ie7 oldie' lang='en'> <![endif]-->" & vbcrlf
emailbody = emailbody & "<!--[if IE 8 ]>    <html class='no-js ie8 oldie' lang='en'> <![endif]-->" & vbcrlf
emailbody = emailbody & "<!--[if IE 9 ]>    <html class='no-js ie9' lang='en'> <![endif]-->" & vbcrlf
emailbody = emailbody & "<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->" & vbcrlf
emailbody = emailbody & "<html class='no-js' lang='en'>" & vbcrlf
emailbody = emailbody & "<!--<![endif]-->" & vbcrlf

emailbody = emailbody & "<head>" & vbcrlf

emailbody = emailbody & "    <meta charset='utf-8'>" & vbcrlf
emailbody = emailbody & "    <meta http-equiv='X-UA-Compatible' content='IE=edge'>" & vbcrlf
emailbody = emailbody & "    <meta name='viewport' content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'>" & vbcrlf
emailbody = emailbody & "    <meta name='description' content=''>" & vbcrlf
emailbody = emailbody & "    <meta name='author' content=''>" & vbcrlf

emailbody = emailbody & "</head>" & vbcrlf

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

If Salesman <> "" Then
	FilterSlsmn1 = Salesman
	FilterSlsmn1 = CInt(FilterSlsmn1)
	FilterSlsmn2 = ""
End If

If SecondarySalesman <> "" Then
	FilterSlsmn2 = SecondarySalesman 
	FilterSlsmn2 = CInt(FilterSlsmn2)
	FilterSlsmn1 = ""
End If

FilterSalesDollars = 100
FilterPercentage = 10

  
emailbody = emailbody & " <style>"

emailbody = emailbody & "   	table.single-salesman-header {"
emailbody = emailbody & "         font-family: verdana, arial, sans-serif;"
emailbody = emailbody & "         font-size: 18px;"
emailbody = emailbody & "         color: #000;"
emailbody = emailbody & "         background-color: #FFF;"
emailbody = emailbody & "         width:100%;"
emailbody = emailbody & "         text-align:left;"
emailbody = emailbody & "         margin:0 auto;"
emailbody = emailbody & "     }"

emailbody = emailbody & "   	table.single-salesman-header td{"
emailbody = emailbody & "   		padding:20px;"
emailbody = emailbody & "     }"

emailbody = emailbody & "   	table.single-salesman {"
emailbody = emailbody & "         font-family: verdana, arial, sans-serif;"
emailbody = emailbody & "         font-size: 11px;"
emailbody = emailbody & "         color: #333333;"
emailbody = emailbody & "         border-width: 1px;"
emailbody = emailbody & "         border-color: #3A3A3A;"
emailbody = emailbody & "         border-collapse: collapse;"
emailbody = emailbody & "     }"
emailbody = emailbody & "     table.single-salesman th {"
emailbody = emailbody & "         border-width: 1px;"
emailbody = emailbody & "         padding: 8px;"
emailbody = emailbody & "         border-style: solid;"
emailbody = emailbody & "         border-color: #d8d8d8;"
emailbody = emailbody & "         background-color: #A4A4A4;"
emailbody = emailbody & "         color: #ffffff;"
emailbody = emailbody & "     }"
emailbody = emailbody & "     table.single-salesman tr:hover td {"
emailbody = emailbody & "         cursor: pointer;"
emailbody = emailbody & "     }"
emailbody = emailbody & "     table.single-salesman tr:nth-child(even) td{"
emailbody = emailbody & "         background-color: #f1f1f1;"
emailbody = emailbody & "     }"
emailbody = emailbody & "     table.single-salesman td {"
emailbody = emailbody & "         border-width: 1px;"
emailbody = emailbody & "         padding: 8px;"
emailbody = emailbody & "         border-style: solid;"
emailbody = emailbody & "         border-color: #d8d8d8;"
emailbody = emailbody & "         background-color: #ffffff;"
emailbody = emailbody & "     }"
    
emailbody = emailbody & " 	 table.single-salesman .vpc-variance-header{"
emailbody = emailbody & " 		background: #D43F3A;"
emailbody = emailbody & " 		color:#fff;"
emailbody = emailbody & " 		text-align:center;"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 	}"
	
emailbody = emailbody & " 	 table.single-salesman .vpc-3pavg-header{"
emailbody = emailbody & " 		background: #F0AD4E;"
emailbody = emailbody & " 		color:#fff;"
emailbody = emailbody & " 		text-align:center;"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 	}"
	
emailbody = emailbody & " 	 table.single-salesman .vpc-lcp-header{"
emailbody = emailbody & " 		background: #337AB7;"
emailbody = emailbody & " 		color:#fff;"
emailbody = emailbody & " 		text-align:center;"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 	}"

emailbody = emailbody & " 	 table.single-salesman .vpc-current-header{"
emailbody = emailbody & " 		background: #5CB85C;"
emailbody = emailbody & " 		color:#fff;"
emailbody = emailbody & " 		text-align:center;"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 	}"

emailbody = emailbody & " 	.gen-info-header{"
emailbody = emailbody & " 		background: #3B579D;"
emailbody = emailbody & " 		color:#fff;"
emailbody = emailbody & " 		text-align:center;"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 	}"

emailbody = emailbody & " 	.negative{"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 		color:red;	"
emailbody = emailbody & " 	}"

emailbody = emailbody & " 	.neutral{"
emailbody = emailbody & " 		font-weight:bold;"
emailbody = emailbody & " 		color:black;"
emailbody = emailbody & " 	}"

emailbody = emailbody & " 	.smaller-header{"
emailbody = emailbody & " 		font-size: 0.8em;"
emailbody = emailbody & " 		vertical-align: top !important;"
emailbody = emailbody & " 		text-align: center;"
emailbody = emailbody & " 	}	"

emailbody = emailbody & " 	.smaller-detail-line{"
emailbody = emailbody & " 		font-size: 0.8em;"
emailbody = emailbody & " 	}	"

emailbody = emailbody & " </style>"


SQL = "SELECT Distinct CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
SQL = SQL & " FROM CustCatPeriodSales "
SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
SQL = SQL & " AND LCPTotSalesAllCats < Total3PPAvgAllCats "
SQL = SQL & " AND Total3PPAvgAllCats - LCPTotSalesAllCats > " & FilterSalesDollars 
SQL = SQL & " AND (CASE WHEN Total3PPAvgAllCats <> 0 THEN (((LCPTotSalesAllCats  - Total3PPAvgAllCats ) / Total3PPAvgAllCats) * 100) * -1 END) >= " & FilterPercentage 
SQL = SQL & " ORDER BY CustNum "

'Response.Write(SQL & "<br>")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
Set rs = cnn8.Execute(SQL)

	
	Do While Not rs.EOF

		ShowThisRecord = True

			
		If ShowThisRecord <> False Then			
		
			PrimarySalesMan =  ""
			SecSalesMan =  ""
			ReferralCode =  ""
			CustomerType =  ""
			SelectedCustomerID = rs("CustNum")
			CustName = GetCustNameByCustNum(SelectedCustomerID)	
			
			'Extra Fields for Filtering
			SQL4 = "SELECT * FROM AR_Customer WHERE CustNum = '" & SelectedCustomerID & "'"
			Set rs4 = Server.CreateObject("ADODB.Recordset")
			rs4.CursorLocation = 3
			Set rs4= cnn8.Execute(SQL4 )

			If Not rs4.Eof Then

				PrimarySalesMan = rs4("Salesman")
				SecSalesMan = rs4("SecondarySalesman")
				ReferralCode = rs4("ReferalCode")
				CustomerType = rs4("CustType")

				'Decide if this record meets the filter criteria
				If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
					If CInt(FilterSlsmn1) <> Cint(rs4("Salesman")) Then ShowThisRecord = False
				End If
				If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
					If CInt(FilterSlsmn2) <> Cint(rs4("SecondarySalesman")) Then ShowThisRecord = False
				End If
				If FilterReferral <> "" And FilterReferral <> "All" Then
					If CInt(FilterReferral) <> Cint(rs4("ReferalCode")) Then ShowThisRecord = False
				End If
				If FilterType <> "" And FilterType <> "All" Then
					If CInt(FilterType) <> Cint(rs4("CustType")) Then ShowThisRecord = False
				End If
			
				Cust_MGP = rs4("ProjGpPerMonth")
				If rs4("ProjSalesPerMonth") <> "" Then Cust_MGPSales = FormatCurrency(rs4("ProjSalesPerMonth"),0) Else Cust_MGPSales =""
							
				MGPTerm = "" 
				' Determine what CCS is going to call it
				If Cust_MGPSales > 0 Then
					If cint(Cust_MGP) = 1 Then
						MGPTerm = "E" 
					Else
						MGPTerm = "C" 
					End If
				Else
					MGPTerm = ""
				End If
				If rs4("ProjSalesPerMonth") <> "" Then Cust_MGPSales = rs4("ProjSalesPerMonth")
				
			Else
				' Customer not found un AR_Customer
				ShowThisRecord = False
			End If

		End If
		
		
		If ShowThisRecord <> False Then
		
			'Get everything we need for the report data
			SQLReportData = "SELECT SUM(TotalSales) AS LCPSales, SUM([3PriorPeriodsTotalSales]) AS ThreePPSales "
			SQLReportData = SQLReportData & ", SUM(PriorPeriod1Sales+PriorPeriod2Sales+PriorPeriod3Sales+PriorPeriod4Sales+PriorPeriod5Sales+PriorPeriod6Sales+ "
			SQLReportData = SQLReportData & " PriorPeriod7Sales+PriorPeriod8Sales+PriorPeriod9Sales+PriorPeriod10Sales+PriorPeriod11Sales+PriorPeriod12Sales) As TwelvePPSales "
			SQLReportData = SQLReportData & " FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & SelectedCustomerID & "' AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
			

			LCPSales = rs("LCPSales")
			ThreePPSales = rs("ThreePPSales")
			TwelvePPSales = rs("TwelvePPSales")
			CurrentPSales = GetCurrent_PostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated)
			LCPvs3PAvgSales = LCPSales - (ThreePPSales/3)
			ImpactDays = (WorkDaysIn3PeriodBasis/3)- WorkDaysInLastClosedPeriod
			DayImpact = ImpactDays  * (LCPSales/WorkDaysInLastClosedPeriod)
			DayImpact = Round(DayImpact,2)
			ADS_LastClosed = (LCPSales/WorkDaysInLastClosedPeriod)
			ADS_3PA = ThreePPSales / (WorkDaysIn3PeriodBasis /3)
			ADS_Variance = ADS_LastClosed -  ADS_3PA 
			LCPvs12PAvgSales = LCPSales - (TwelvePPSales/12)
			If LCPvs12PAvgSales <> 0 Then LCPvs12PAvgPercent = ((LCPSales - LCPvs12PAvgSales) / LCPvs12PAvgSales)  * 100 Else LCPvs12PAvgPercent = 0
			SamePLYSales = TotalTPLYAllCats(PeriodSeqBeingEvaluated,SelectedCustomerID)
			ThreePPAvgSales = ThreePPSales / 3
			TwelvePPAvgSales = TwelvePPSales / 12
			If ThreePPAvgSales <> 0 Then LCPvs3PAvgPercent = ((LCPSales - ThreePPAvgSales ) / ThreePPAvgSales )  * 100  Else LCPvs3PAvgPercent = 0
			If MGPTerm <> "" Then LCPvsMxS = LCPSales - Cust_MGPSales
			If MGPTerm <> "" Then ThreePAvgVMxS = ThreePPAvgSales - Cust_MGPSales
			If MGPTerm <> "" Then TwelvePAvgVMxS = TwelvePPAvgSales - Cust_MGPSales
			If MGPTerm <> "" Then CurrentVMxS = CurrentPSales - Cust_MGPSales
			'ROI***********
			TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(SelectedCustomerID)
			'If CustHasEquipment(SelectedCustomerID) Then
			If TotalEquipmentValue > 0 Then	
				'LCPGP = LCPSales - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,SelectedCustomerID)
				LCPGP = LCPSales - rs("TotalCostAllCats")
				ThreePAvgGP = ThreePPAvgSales - ( TotalCostByPeriodSeqPrior3P(PeriodSeqBeingEvaluated,SelectedCustomerID) / 3 )
				If LCPGP <> 0 Then ROI = TotalEquipmentValue/LCPGP Else ROI = ""
				If ThreePAvgGP <> 0 Then ROI3P = TotalEquipmentValue/ThreePAvgGP Else ROI3P = ""
			End If

			' HERE ARE THE RULES
			'1. If Current (adjusted for days) => 3Pavg or 12Pavg - Dont show
			If CurrentPSales >= ThreePPAvgSales OR CurrentPSales >= TwelvePPAvgSales Then  'If the current is already greater or equal, we don't need to adjust for days, we're already ok
				ShowThisRecord = False
			Else  'We need to adjust for days & fiure it out
				ForecastedCurrent = (CurrentPSales / WorkDaysSoFar) * WorkDaysInCurrentPeriod 
				If ForecastedCurrent >= ThreePPAvgSales OR ForecastedCurrent >= TwelvePPAvgSales Then ShowThisRecord = False
			End If
				
			'2. If LCP => 12pAVG - Dont Show
			If LCPSales >= TwelvePPAvgSales Then ShowThisRecord = False
			
			'3. If LCP => SPLY - Dont Show
			If LCPSales >= SamePLYSales Then ShowThisRecord = False
			
				
			'4. If 3PROI > 10 - Override anything else and Show
			If Not Isnull(ROI3P) Then
				If IsNumeric(ROI3P) Then
					If ROI3P > 10 Then ShowThisRecord = True
				End If
			End If


			If ShowThisRecord <> False Then TotalCustsReported = TotalCustsReported + 1

		End If
		
		rs.movenext
			
	Loop
	

rs.Close	

emailbody = emailbody & " <table class='single-salesman-header'>"

emailbody = emailbody & " 	<tr>"
emailbody = emailbody & " 		<td>"

If Salesman <> "" Then
	emailbody = emailbody & "Hi " & GetUserDisplayNameByUserNo(GetUserNoBySalesPersonNo(FilterSlsmn1)) & ","
End If
If SecondarySalesman <> "" Then
	emailbody = emailbody & "Hi " & GetUserDisplayNameByUserNo(GetUserNoBySalesPersonNo(FilterSlsmn2)) & ","
End If

emailbody = emailbody & " 		</td>"
emailbody = emailbody & " 	</tr>"
	
emailbody = emailbody & " 	<tr>"
emailbody = emailbody & " 		<td>"
emailbody = emailbody & " 			Your " & GetTerm("customer") & " analysis report has been prepared and is ready for you to view."
emailbody = emailbody & " 			There are " & TotalCustsReported & " " & GetTerm("customers") & " that need your attention."
emailbody = emailbody & " 		</td>"
emailbody = emailbody & " 	</tr>"
emailbody = emailbody & " 	<tr>"
emailbody = emailbody & " 		<td>"

'************
' Quick Login
'************

destination = "bizintel/CustAnalSum_1.asp"

If MUV_READ("ClientID") = "1071" or  MUV_READ("SERNO") = "1071d" Then
	linkvar = "ql-CCS.asp" 
Else
	linkvar = "ql.asp"
End IF

emailbody = emailbody & "<a href='" & baseURL & linkvar & "?"

If Salesman <> "" Then 
	emailbody = emailbody & "c=" & MUV_READ("ClientID") & "&u=" & GetUserNoBySalesPersonNo(FilterSlsmn1) & "&d=" & destination & "-qlSls=" & FilterSlsmn1
End If
If SecondarySalesman <> "" Then
	emailbody = emailbody & "c=" & MUV_READ("ClientID") & "&u=" & GetUserNoBySalesPersonNo(FilterSlsmn2) & "&d=" & destination & "-qlSls2=" & FilterSlsmn2
End If

emailbody = emailbody &  "'>"
	
emailbody = emailbody & " 			Click here to log in to MDS Insight and begin reviewing your report"

emailbody = emailbody & "</a>"
emailbody = emailbody & " 		</td>"
emailbody = emailbody & " 	</tr>"

For z = 1 to 2
	emailbody = emailbody & "<tr><td>&nbsp;</td></tr>"
Next	

emailbody = emailbody & "<tr><td class='smaller-detail-line'>(Insight CID: " & MUV_READ("ClientID") & ")</td></tr>"

emailbody = emailbody & " </table>"

%>