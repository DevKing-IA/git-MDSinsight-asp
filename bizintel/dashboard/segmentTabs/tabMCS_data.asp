<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../../inc/InSightFuncs_AR_AP.asp"-->


<%
	Segment = Request.QueryString("p")

	
	ShowPercentageColumns = False

		
	PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
	PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()
	
	WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
	WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
	WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
	WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
	WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1


	JSON=""

	
	Select Case MUV_READ("LOHVAR")
		Case "Secondary"

			SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
			SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
			SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "	
			SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
			SQL = SQL & " AND  CustCatPeriodSales_ReportData.SecondarySalesman = " & Segment 
			SQL = SQL & " AND  AR_Customer.MonthlyContractedSalesDollars IS NOT NULL"
	
		Case "Primary"

			SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
			SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
			SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "	
			SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
			SQL = SQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman = " & Segment 
			SQL = SQL & " AND  AR_Customer.MonthlyContractedSalesDollars IS NOT NULL"
	
		Case "CustType"

			SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
			SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
			SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "	
			SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
			SQL = SQL & " AND  CustCatPeriodSales_ReportData.CustType = " & Segment 
			SQL = SQL & " AND  AR_Customer.MonthlyContractedSalesDollars IS NOT NULL"
	
	End Select	
	
'	Response.write(SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.ConnectionTimeout = 120
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)
	

		GrandTotLCPvs3PAvgSales = 0
				
		Do While Not rs.EOF

			ShowThisRecord = True

				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
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
					SecondarySalesMan = rs4("SecondarySalesman")
					ReferralCode = rs4("ReferalCode")
					CustomerType = GetCustTypeByCode(rs4("CustType"))
					CustomerType = Replace(CustomerType,"CLIENT","")


					MonthlyContractedSalesDollars = rs4("MonthlyContractedSalesDollars")
					
				Else
					' Customer not found un AR_Customer
					ShowThisRecord = False
				End If

			End If
			
			
			If ShowThisRecord <> False Then
			
				
				PP1Sales = 0
				PP2Sales = 0
				
				'Now quick get the Prior Period 1 and Prior Period 2 Sales
				Set rs35 = Server.CreateObject("ADODB.Recordset")
				rs35.CursorLocation = 3
				SQL35 = "SELECT Sum(PriorPeriod1Sales) As PP1, Sum(PriorPeriod2Sales) As PP2 "
				SQL35 = SQL35 & " FROM CustCatPeriodSales_ReportData "
				SQL35 = SQL35 & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL35 = SQL35 & " AND  CustNum = '" & SelectedCustomerID & "'"
				Set rs35= cnn8.Execute(SQL35)
				If Not rs35.EOF Then
					PP1Sales = rs35("PP1")
					PP2Sales = rs35("PP2")
				End If

				LCPSales = rs("LCPSales")
				If Not IsNumeric(LCPSales) Then LCPSales = 0
				ThreePPSales = rs("ThreePPSales")
				TwelvePPSales = rs("TwelvePPSales")
				CurrentPSales = GetCurrent_PostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated)
				LCPvs3PAvgSales = LCPSales - (ThreePPSales/3)
				If Not IsNumeric(LCPvs3PAvgSales) Then LCPvs3PAvgSales = 0

				ImpactDays = (WorkDaysIn3PeriodBasis/3)- WorkDaysInLastClosedPeriod
				DayImpact = ImpactDays  * (LCPSales/WorkDaysInLastClosedPeriod)
				DayImpact = Round(DayImpact,2)
				ADS_LastClosed = (LCPSales/WorkDaysInLastClosedPeriod)
				ADS_3PA = ThreePPSales / (WorkDaysIn3PeriodBasis /3)
				ADS_Variance = ADS_LastClosed -  ADS_3PA 
				If Not IsNumeric(ADS_Variance) Then ADS_Variance = 0
				LCPvs12PAvgSales = LCPSales - (TwelvePPSales/12)
				If Not IsNumeric(LCPvs12PAvgSales) Then LCPvs12PAvgSales = 0
				If LCPvs12PAvgSales <> 0 Then LCPvs12PAvgPercent = ((LCPSales - LCPvs12PAvgSales) / LCPvs12PAvgSales)  * 100 Else LCPvs12PAvgPercent = 0
				SamePLYSales = TotalTPLYAllCats(PeriodSeqBeingEvaluated,SelectedCustomerID)
				If Not IsNumeric(SamePLYSales) Then SamePLYSales = 0
				ThreePPAvgSales = ThreePPSales / 3
				TwelvePPAvgSales = TwelvePPSales / 12
				If ThreePPAvgSales <> 0 Then LCPvs3PAvgPercent = ((LCPSales - ThreePPAvgSales ) / ThreePPAvgSales )  * 100  Else LCPvs3PAvgPercent = 0
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

				If Not IsNumeric(ThreePPAvgSales) Then ThreePPAvgSales = 0
				If Not IsNumeric(TwelvePPAvgSales) Then TwelvePPAvgSales = 0

				
			
				
				If ShowThisRecord <> False Then
				
					TotalCustsReported = TotalCustsReported + 1
					
					GrandTotLCPvs3PAvgSales = GrandTotLCPvs3PAvgSales + LCPvs3PAvgSales					
					IF LEN(JSON)>0 Then
						JSON=JSON+","
					END If
					JSON=JSON+"{"
					JSON=JSON & """SelectedCustomerID"":""" & SelectedCustomerID & """"
					JSON=JSON+","
					JSON=JSON & """CustName"":""" & CustName & """"
					JSON=JSON+","
					JSON=JSON & """LCPvs3PAvgSales"":""" & FormatCurrency(LCPvs3PAvgSales,0,-2,0) & """"
					JSON=JSON+","
					JSON=JSON & """LCPvs3PAvgPercent"":""" & FormatNumber(LCPvs3PAvgPercent,0) & """"
					JSON=JSON+","
					JSON=JSON & """DayImpact"":""" & FormatCurrency(DayImpact,0) & """"
					JSON=JSON+","
					JSON=JSON & """ADS_Variance"":""" & FormatCurrency(ADS_Variance,0) & """"
					JSON=JSON+","
					JSON=JSON & """LCPvs12PAvgSales"":""" & FormatCurrency(LCPvs12PAvgSales,0) & """"
					JSON=JSON+","
					JSON=JSON & """LCPvs12PAvgPercent"":""" & FormatNumber(LCPvs12PAvgPercent,0)  & """"
					JSON=JSON+","
					JSON=JSON & """PP1Sales"":""" & FormatCurrency(PP1Sales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """PP2Sales"":""" & FormatCurrency(PP2Sales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """LCPSales"":""" & FormatCurrency(LCPSales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """ThreePPAvgSales"":""" & FormatCurrency(ThreePPAvgSales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """TwelvePPAvgSales"":""" & FormatCurrency(TwelvePPAvgSales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """CurrentPSales"":""" & FormatCurrency(CurrentPSales,0,-2,0)  & """"
					JSON=JSON+","
					JSON=JSON & """SamePLYSales"":""" & FormatCurrency(SamePLYSales,0,-2,0)  & """"
					JSON=JSON+","

					If Not IsNull(MonthlyContractedSalesDollars) Then
						JSON=JSON & """MCS"":""" &  FormatCurrency(MonthlyContractedSalesDollars,0)  & """"
						JSON=JSON+","
						
					Else
						JSON=JSON & """MCS"":""0"""
						JSON=JSON+","
		
					End If
					

					If Not IsNull(MonthlyContractedSalesDollars) Then
						
						JSON=JSON & """LCPvsMCS"":""" &  FormatCurrency(LCPSales-MonthlyContractedSalesDollars,0,-2,0)  & """"
						JSON=JSON+","
						
					Else
						JSON=JSON & """LCPvsMCS"":"""""
						JSON=JSON+","
						
					End If

					

					If Not IsNull(MonthlyContractedSalesDollars) Then
						JSON=JSON & """3PavgvsMCS"":""" &  FormatCurrency(ThreePPAvgSales-MonthlyContractedSalesDollars,0,-2,0)  & """"
						JSON=JSON+","
						
					
					Else
						JSON=JSON & """3PavgvsMCS"":"""""
						JSON=JSON+","
						
					End If
					
										

					If Not IsNull(MonthlyContractedSalesDollars) Then
						
						JSON=JSON & """12PavgvsMCS"":""" &  FormatCurrency(TwelvePPAvgSales-MonthlyContractedSalesDollars,0,-2,0)  & """"
						JSON=JSON+","
						
					
						Else
						JSON=JSON & """12PavgvsMCS"":"""""
						JSON=JSON+","
						
						
						
					End If
					



					If Not IsNull(MonthlyContractedSalesDollars) Then
						JSON=JSON & """CurrentvsMCS"":""" &  FormatCurrency(CurrentPSales-MonthlyContractedSalesDollars,0,-2,0)  & """"
						JSON=JSON+","
						
					
						Else
						JSON=JSON & """CurrentvsMCS"":"""""
						JSON=JSON+","
						
					End If
	
					If TotalEquipmentValue > 0 Then	
						If IsNumeric(ROI) Then
								JSON=JSON & """LCP_ROI"":""" &   FormatNumber(ROI,1)  & """"
								JSON=JSON+","
							Else

								JSON=JSON & """LCP_ROI"":""No Sales"""
								JSON=JSON+","
						End If
						If IsNumeric(ROI3P) Then
								JSON=JSON & """PavgROI"":""" & FormatNumber(ROI3P,1) & """"
								JSON=JSON+","

							Else
								JSON=JSON & """PavgROI"":"""""
								JSON=JSON+","
								'Response.Write("<td>&nbsp;</td>")
						End If
						' Write equipment value regardless of ROI
						JSON=JSON & """TotalEquipmentValue"":""" & FormatCurrency(TotalEquipmentValue,0) & """"
						JSON=JSON+","
					Else
						JSON=JSON & """LCP_ROI"":"""""
						JSON=JSON+","
						JSON=JSON & """PavgROI"":"""""
						JSON=JSON+","
						JSON=JSON & """TotalEquipmentValue"":"""""
						JSON=JSON+","
					End If
	
	
					' General info
					PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
					SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesMan)
					
					Select Case MUV_READ("LOHVAR")
							Case "Secondary"
							    If Instr(PrimarySalesPerson ," ") <> 0 Then
									JSON=JSON & """PrimarySalesPerson"":""" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & """"
									JSON=JSON+","
								Else
									JSON=JSON & """PrimarySalesPerson"":""" & PrimarySalesPerson & """"
									JSON=JSON+","
								End If
							Case "Primary"
							    If Instr(SecondarySalesPerson ," ") <> 0 Then
									JSON=JSON & """SecondarySalesPerson"":""" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson ," ")+1) & """"
									JSON=JSON+","
								Else
									JSON=JSON & """SecondarySalesPerson"":""" & SecondarySalesPerson & """"
									JSON=JSON+","
								End If
							Case "CustType"
							    If Instr(SecondarySalesPerson ," ") <> 0 Then
									JSON=JSON & """SecondarySalesPerson"":""" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson ," ")+1) & """"
									JSON=JSON+","
								Else
									JSON=JSON & """SecondarySalesPerson"":""" & SecondarySalesPerson & """"
									JSON=JSON+","
								End If
								
					End Select	
					JSON=JSON & """CustomerType"":""" & CustomerType & """"
					JSON=JSON+","
					JSON=JSON & """CustomerNotes"":""" & UserHasAnyUnviewedNotes(SelectedCustomerID) & """"
					JSON=JSON+","
					JSON=JSON & """rules"":""" & "123abc" & """"

	                JSON=JSON & "}"
				    'Response.Write("</tr>")
			    
			    End If

			End If
			
			rs.movenext
				
		Loop
		'retData="{""orderby"":""" & orderValue & """,""draw"": " & CLng(Request.QueryString("draw")) & ",""recordsTotal"": " & nRecordCount & ",""recordsFiltered"": " & nRecordCount & ",""data"": [" & JSONdata & "],""byRegionData"":"+GetQtyCustByRegion()+"}"
		JSON="{""data"":[" & JSON & "]}"
		
		Response.AddHeader "Content-Type", "application/json"
		response.write JSON

%>

