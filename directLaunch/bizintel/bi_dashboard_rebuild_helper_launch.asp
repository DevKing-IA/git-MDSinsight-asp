<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 
<%

IncludeUnposted = True



'Response.Buffer = True  <-----
'Response.Expires = 0  <-----	These lines commented purposely. They keep the page from close when launched automatically. Can't use them.
'Response.Clear  <-----


'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/bizintel/bi_dashboard_rebuild_helper_launch.asp?runlevel=run_now
Server.ScriptTimeout = 75000

Dim EntryThread


'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 

SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 and ClientKey='1071d'"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
	
		'To begin with, see if this client uses the Biz Intel 
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("biModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then

												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				

				'**************************************************************
				'Get next Entry Thread for use in the SC_AuditLogDLaunch table
				On Error Goto 0
				Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
				cnnAuditLog.open MUV_READ("ClientCnnString") 
				Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
				rsAuditLog.CursorLocation = 3 
				Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
				If Not rsAuditLog.EOF Then 
					If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
				Else
					EntryThread = 1
				End If
				set rsAuditLog = nothing
				cnnAuditLog.close
				set cnnAuditLog = nothing
					
	
					CreateAuditLogEntry "BI Dashboard Rebuild Helper Launch","BI Dashboard Rebuild Helper Launch","Minor",0,"BI Dashboard Rebuild Helper Launch ran."					
	
					WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"
	
					
					If MUV_READ("cnnStatus") = "OK" Then ' else it loops
					
						Response.Write("blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah <br>")
						
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''

''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''
'S E C O N D A R Y  S A L E S M A N
''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''

						Set rs = Server.CreateObject("ADODB.Recordset")
						Set rs2 = Server.CreateObject("ADODB.Recordset")

						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.open MUV_READ("ClientCnnString") 

						SQL = "DELETE FROM BI_Dashboard WHERE Segment = 'SecondarySalesman'"
						Set rs = cnn8.Execute(SQL)

						PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
						PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()
	
						WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
						WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
						WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) 
						WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
						WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1
						WorkDaysInProjectionBasis =  (NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -2), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1) + WorkDaysInLastClosedPeriod

						LCP_Display_MonthArr = Split(GetLastClosedPeriodAndYear(),"-")
						LCP_Display_Month = Trim(LCP_Display_MonthArr(0))
						LCP_Display_Year = Trim(LCP_Display_MonthArr(1))
	
						LCP_Display_Var = "P" & LCP_Display_Month & "/" & LCP_Display_Year
	
						SQL = "SELECT BeginDate As SPLYBeginDate, EndDate As SPLYEndDate FROM BillingPeriodHistory WHERE Period = " & LCP_Display_Month & " AND Year = " & LCP_Display_Year - 1
						Set rs = cnn8.Execute(SQL)
						SPLYBeginDate = rs("SPLYBeginDate")
						SPLYEndDate = rs("SPLYEndDate")
						WorkDaysInSPLYPeriodBasis =  NumberOfWorkDays(SPLYBeginDate, SPLYEndDate)


						SQL = "SELECT BeginDate As CPLYBeginDate, EndDate As CPLYEndDate FROM BillingPeriodHistory WHERE Period = " & LCP_Display_Month+1 & " AND Year = " & LCP_Display_Year - 1
						Set rs = cnn8.Execute(SQL)
						CPLYBeginDate = rs("CPLYBeginDate")
						CPLYEndDate = rs("CPLYEndDate")
						WorkDaysInCurrentPLY =  NumberOfWorkDays(CPLYBeginDate , CPLYEndDate )

						SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
						SQL = SQL & " FROM CustCatPeriodSales "
						SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
						Set rs = cnn8.Execute(SQL)
	
						TotSalesCPLY = rs("CPLYTotalSales")
						TotRentalsCPLY = rs("CPLYTotalRentals")

						'Current Period Last Year
						'See if it has a decimal > 5
						If WorkDaysInCurrentPLY  - Int(WorkDaysInCurrentPLY) < .5 Then
							WorkDaysInCurrentPLY = Int(WorkDaysInCurrentPLY)
						Else
							WorkDaysInCurrentPLY = Int(WorkDaysInCurrentPLY) + .5
						End If

						Tot_CPLYPADS = TotSalesCPLY / WorkDaysInCurrentPLY 
					
						WD_P3PADS = WorkDaysIn3PeriodBasis / 3
						'See if it has a decimal > 5
						If WD_P3PADS - Int(WD_P3PADS) < .5 Then
							WD_P3PADS = Int(WD_P3PADS)
						Else
							WD_P3PADS = Int(WD_P3PADS) + .5
						End If
						Tot_P3PADS = Tot3PAvgHeader / WD_P3PADS


						WD_P12PADS = WorkDaysIn12PeriodBasis / 12
						'See if it has a decimal > 5
						If WD_P12PADS - Int(WD_P12PADS ) < .5 Then
							WD_P12PADS = Int(WD_P12PADS )
						Else
							WD_P12PADS = Int(WD_P12PADS ) + .5
						End If
						Tot_P12PADS = Tot12PAvgHeader / WD_P12PADS
	
	
						WD_SPLYPADS = WorkDaysInSPLYPeriodBasis 
						'See if it has a decimal > 5
						If WD_SPLYPADS - Int(WD_SPLYPADS ) < .5 Then
							WD_SPLYPADS = Int(WD_P12PADS )
						Else
							WD_SPLYPADS = Int(WD_SPLYPADS ) + .5
						End If
						Tot_SPLYPADS = SPLYTotalSales  / WD_SPLYPADS
	
						'Current Period
						WD_CurrentSoFar = WorkDaysSoFar 
						WD_CurrentPeriod = WorkDaysInCurrentPeriod  
						'See if it has a decimal > 5
						If WD_SPLYPADS - Int(WD_SPLYPADS ) < .5 Then
							WD_SPLYPADS = Int(WD_P12PADS )
						Else
							WD_SPLYPADS = Int(WD_SPLYPADS ) + .5
						End If
							
					
						TotProductSalesSls2 = 0
						Tot3PAvgSls2 = 0
						TotDollarDiff =0
						TotalNegDiff = 0
						Tot12PAvgSls2 = 0
						TotDollarDiff12 = 0
						TotalNegDiff12 = 0
					
						GrandTotalLCPSales = 0 : GrandTotal3PAvgSales = 0 : GrantTotal12PAvgSales = 0 : GrandTotalSPLYSales = 0
						GrandTotalLCPADS = 0 : GrandTotal3PPADS = 0 : GrandTotal12PPADS = 0 : GrandTotalSPLYADS = 0 : GrandTotalCPADS = 0 : GrandTotalCPLYADS = 0
						GrandTotalLCPvs3PAvgADS = 0 : GrandTotalLCPvs12PAvg = 0 : GrandTotalLCPvsSPLY = 0 : GrandTotalCPvsvs3PAvgADS = 0 : GrandTotalCPLYvsvs3PAvgADS = 0
					
						LeftOverSLs2 = ""

						SQL = "SELECT SecondarySalesman "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
						SQL = SQL & " GROUP BY SecondarySalesman"
						SQL = SQL & " EXCEPT "
						SQL = SQL & " SELECT SecondarySalesman "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
						SQL = SQL & " GROUP BY SecondarySalesman"
					
					'Response.Write(SQL)
				
					Set rs = cnn8.Execute(SQL)
					If Not rs.EOF Then
						Do While Not rs.EOF
								LeftOverSLs2 = LeftOverSLs2 & rs("SecondarySalesman") & ","
							rs.MoveNext
						Loop
					End IF

					If Right(LeftOverSLs2,1)="," Then LeftOverSLs2 = Left(LeftOverSLs2 ,len(LeftOverSLs2 )-1)
					
				'Response.Write(SQL&"<br>")	
					Set rs = cnn8.Execute(SQL)


	
					SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS TotProductSales ,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
					SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
					SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
					SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
					SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
					SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "	
					SQL = SQL & ",SecondarySalesman"
					SQL = SQL & " FROM CustCatPeriodSales_ReportData "
					SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					SQL = SQL & " GROUP BY SecondarySalesman"
					SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"
				
				'Response.Write(SQL&"<br>")	
					Set rs = cnn8.Execute(SQL)
					
					If not rs.EOF Then

	
						ChartElementNumber = 1
						ChartDataSls2 = ""
						ChartRemainder = 100
						NextPeriodProj = 0

									
						Do
						
							P3PAvgProductSales = rs("Tot3PPAvg") - rs("Tot3PPAvgRentals")
							P12PAvgProductSales = rs("Tot12PPAvg") - rs("Tot12PPAvgRentals")
							SPLYProductsSales = 0
							SPLYTotalRentals =  0
							LCPADS = 0
							P3PADS = 0
							P12PADS = 0
							SPLYADS = 0
							CPADS = 0
							CPLYADS = 0

							'No link if lcp + 3pp < $1
							TotalToEval = rs("TotProductSales") + P3PAvgProductSales 
							If Not Isnumeric(TotalToEval) Then TotalToEval = 0
			
							Sls22Find = rs("SecondarySalesman")
							
							'Now get all the SPLY Numbers
							SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
							SQL2 = SQL2 & " FROM CustCatPeriodSales "
							SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL2 = SQL2 & " AND SecondarySalesman = '" & Sls22Find & "' "
							
							Set rs2 = cnn8.Execute(SQL2)
							If not rs2.EOF Then
								SPLYProductsSales = rs2("SPLYTotalSales")
								SPLYTotalRentals = rs2("SPLYTotalRentals")
							End If

							If IncludeUnposted = True Then
								RentalsHolder = GetCurrentPeriod_PostedRentalsSls2(Sls22Find) + GetCurrentPeriod_UnPostedRentalsSls2(Sls22Find)
								ProdSalesHolder = (GetCurrentPeriod_PostedTotalSls2(Sls22Find) + GetCurrentPeriod_UnPostedTotalSls2(Sls22Find)) - RentalsHolder 
							Else
								RentalsHolder = GetCurrentPeriod_PostedRentalsSls2(Sls22Find)
								ProdSalesHolder = GetCurrentPeriod_PostedTotalSls2(Sls22Find) - RentalsHolder 
							End If
							
			
							CP = ProdSalesHolder 
				
							SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
							SQL = SQL & " FROM CustCatPeriodSales "
							SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL = SQL & " AND SecondarySalesman = '" & Sls22Find & "'"
							Set rs2 = cnn8.Execute(SQL)
							
							TotSalesCPLY = rs2("CPLYTotalSales")
							TotRentalsCPLY = rs2("CPLYTotalRentals")

							'***********************************
							'***********************************
							' Here is all the ADS variance stuff
							'***********************************
							'***********************************
							LCPADS = Round(rs("TotProductSales"),0)/WorkDaysInLastClosedPeriod
							P3PADS = Round(P3PAvgProductSales,0)/WD_P3PADS
							P12PADS = Round(P12PAvgProductSales,0)/WD_P12PADS
							SPLYADS = Round(SPLYProductsSales,0)/WD_SPLYPADS
							CPADS = Round(CP,0)/WD_CurrentSoFar
							CPLYADS = TotSalesCPLY/WorkDaysInCurrentPLY
				
			
							'Sales
							GrandTotalLCPSales = GrandTotalLCPSales + rs("TotProductSales")
							GrandTotal3PAvgSales = GrandTotal3PAvgSales + P3PAvgProductSales 
							GrantTotal12PAvgSales = GrantTotal12PAvgSales + P12PAvgProductSales 
							GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales  
							
							'ADS
							GrandTotalLCPADS = GrandTotalLCPADS + LCPADS 
							GrandTotal3PPADS = GrandTotal3PPADS + P3PADS 
							GrandTotal12PPADS = GrandTotal12PPADS + P12PADS 
							GrandTotalSPLYADS = GrandTotalSPLYADS + SPLYADS 
							GrandTotalCPADS = GrandTotalCPADS + CPADS 
							GrandTotalCPLYADS = GrandTotalCPLYADS + CPLYADS 
							
							'ADS Variance
							GrandTotalLCPvs3PAvgADS = GrandTotalLCPADS - GrandTotal3PPADS 
							GrandTotalLCPvs12PAvg = GrandTotalLCPADS - GrandTotal12PPADS 
							GrandTotalLCPvsSPLY = GrandTotalLCPADS - GrandTotalSPLYADS 
							GrandTotalCPvsvs3PAvgADS = GrandTotalCPADS - GrandTotal3PPADS 
							GrandTotalCPLYvsvs3PAvgADS = GrandTotalCPLYADS - GrandTotal3PPADS 
							
							
							If rs("SecondarySalesman") <> 0 Then
								SQL2 = "INSERT INTO BI_Dashboard ( "
								SQL2 = SQL2 & " SecondarySalesmanNumber "
								SQL2 = SQL2 & " , SecondarySalesmanName "
								SQL2 = SQL2 & " , SalesLCP "
								SQL2 = SQL2 & " , Sales3PPAvg "
								SQL2 = SQL2 & " , Sales12PPAvg "
								SQL2 = SQL2 & " , SalesSPLY "
								SQL2 = SQL2 & " , ADSLCP "
								SQL2 = SQL2 & " , ADS3PPAvg "
								SQL2 = SQL2 & " , ADS12PPAvg "
								SQL2 = SQL2 & " , ADSSPLY "
								SQL2 = SQL2 & " , ADSCP "
								SQL2 = SQL2 & " , ADSCPLY "
								SQL2 = SQL2 & " , VARLCPv3ppAvg "
								SQL2 = SQL2 & " , VARLCPv12PPAvg "
								SQL2 = SQL2 & " , VARLCPvSPLY "
								SQL2 = SQL2 & " , VARCPv3PPAvg "
								SQL2 = SQL2 & " , VARCPvCPLY"
								SQL2 = SQL2 & " , Segment) VALUES ("
								
								SQL2 = SQL2 & "'" & rs("SecondarySalesman") & "'"
								SQL2 = SQL2 & ",'" & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "'"
								SQL2 = SQL2 & "," & rs("TotProductSales")
								SQL2 = SQL2 & "," & P3PAvgProductSales
								SQL2 = SQL2 & "," & P12PAvgProductSales
								SQL2 = SQL2 & "," & SPLYProductsSales
								SQL2 = SQL2 & "," & LCPADS 
								SQL2 = SQL2 & "," & P3PADS 
								SQL2 = SQL2 & "," & P12PADS 
								SQL2 = SQL2 & "," & SPLYADS 
								SQL2 = SQL2 & "," & CPADS 
								SQL2 = SQL2 & "," & CPLYADS 
								SQL2 = SQL2 & "," & LCPADS - P3PADS
								SQL2 = SQL2 & "," & LCPADS - P12PADS
								SQL2 = SQL2 & "," & LCPADS - SPLYADS				
								SQL2 = SQL2 & "," & CPADS - P3PADS
								SQL2 = SQL2 & "," & CPADS - CPLYADS 
								SQL2 = SQL2 & ", 'SecondarySalesman')"
								Set rs2 = cnn8.Execute(SQL2)
							End If
					
				rs.movenext
			Loop until rs.eof
		End If



		'***********
		'***********
		' LEFT OVERS
		'***********
		'***********
		If LeftOverSLs2  <> "" Then
		
	      	'This part is a little crazy but now we have to ddo the left over salesman2's
			'Now get all the SPLY Numbers for the leftovers
			SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
			SQL2 = SQL2 & " ,SecondarySalesman "
			SQL2 = SQL2 & " FROM CustCatPeriodSales "
			SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
			SQL2 = SQL2 & " AND SecondarySalesman IN (" & LeftOverSLs2 & ") GROUP BY SecondarySalesman "

	
			Set rs = cnn8.Execute(SQL2)
			If Not rs.EOF Then
				Do While Not rs.EOF
				
					SPLYProductsSales = rs2("SPLYTotalSales")
					SPLYTotalRentals = rs2("SPLYTotalRentals")
				
					If SPLYProductsSales <> 0 Then 
					
						GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales 			
						GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
	
					End If
								
					rs.MoveNext
				Loop
			End IF
		End If
		'***************
		'***************
		' END LEFT OVERS
		'***************
		'***************

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  S E C O N D A R Y  S A L E S M A N
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'P R I M A R Y  S A L E S M A N
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

						SQL = "DELETE FROM BI_Dashboard WHERE Segment = 'PrimarySalesman'"
						Set rs = cnn8.Execute(SQL)


					 	'Get all Slsmn 1
					 	
						TotProductSalesSls1 = 0
						Tot3PAvgSls1 = 0
						TotDollarDiff =0
						TotalNegDiff = 0
						Tot12PAvgSls1 = 0
						TotDollarDiff12 = 0
						TotalNegDiff12 = 0
					
						GrandTotalLCPSales = 0 : GrandTotal3PAvgSales = 0 : GrantTotal12PAvgSales = 0 : GrandTotalSPLYSales = 0
						GrandTotalLCPADS = 0 : GrandTotal3PPADS = 0 : GrandTotal12PPADS = 0 : GrandTotalSPLYADS = 0 : GrandTotalCPADS = 0 : GrandTotalCPLYADS = 0
						GrandTotalLCPvs3PAvgADS = 0 : GrandTotalLCPvs12PAvg = 0 : GrandTotalLCPvsSPLY = 0 : GrandTotalCPvsvs3PAvgADS = 0 : GrandTotalCPLYvsvs3PAvgADS = 0

						LeftOverSLs1 = ""
					
						SQL = "SELECT PrimarySalesman "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
						SQL = SQL & " GROUP BY PrimarySalesman"
						SQL = SQL & " EXCEPT "
						SQL = SQL & " SELECT PrimarySalesman "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
						SQL = SQL & " GROUP BY PrimarySalesman"

					
						Set rs = cnn8.Execute(SQL)
						If Not rs.EOF Then
							Do While Not rs.EOF
									LeftOverSLs1 = LeftOverSLs1 & rs("PrimarySalesman") & ","
								rs.MoveNext
							Loop
						End IF
					
						If Right(LeftOverSLs1,1)="," Then LeftOverSLs1 = Left(LeftOverSLs1 ,len(LeftOverSLs1 )-1)
						
					'Response.Write(SQL&"<br>")	
						Set rs = cnn8.Execute(SQL)


	
						SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS TotProductSales ,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
						SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
						SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
						SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "	
						SQL = SQL & ",PrimarySalesman"
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					
					
						SQL = SQL & " GROUP BY PrimarySalesman"
						SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"
					
					'Response.Write(SQL&"<br>")	
						Set rs = cnn8.Execute(SQL)
	
						If not rs.EOF Then

					
							ChartElementNumber = 1
							ChartDataSls1 = ""
							ChartRemainder = 100
							NextPeriodProj = 0

									
							Do
							
								P3PAvgProductSales = rs("Tot3PPAvg") - rs("Tot3PPAvgRentals")
								P12PAvgProductSales = rs("Tot12PPAvg") - rs("Tot12PPAvgRentals")
								SPLYProductsSales = 0
								SPLYTotalRentals =  0
								LCPADS = 0
								P3PADS = 0
								P12PADS = 0
								SPLYADS = 0
								CPADS = 0
								CPLYADS = 0

								'No link if lcp + 3pp < $1
								TotalToEval = rs("TotProductSales") + P3PAvgProductSales 
								If Not Isnumeric(TotalToEval) Then TotalToEval = 0
				
								Sls22Find = rs("PrimarySalesman")
								
								'Now get all the SPLY Numbers
								SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
								SQL2 = SQL2 & " FROM CustCatPeriodSales "
								SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
								SQL2 = SQL2 & " AND PrimarySalesman = '" & Sls22Find & "' "
								
								Set rs2 = cnn8.Execute(SQL2)
								If not rs2.EOF Then
									SPLYProductsSales = rs2("SPLYTotalSales")
									SPLYTotalRentals = rs2("SPLYTotalRentals")
								End If

								If IncludeUnposted = True Then
									RentalsHolder = GetCurrentPeriod_PostedRentalsSls(Sls22Find) + GetCurrentPeriod_UnPostedRentalsSls(Sls22Find)
									ProdSalesHolder = (GetCurrentPeriod_PostedTotalSls(Sls22Find) + GetCurrentPeriod_UnPostedTotalSls(Sls22Find)) - RentalsHolder 
								Else
									RentalsHolder = GetCurrentPeriod_PostedRentalsSls(Sls22Find) 
									ProdSalesHolder = GetCurrentPeriod_PostedTotalSls(Sls22Find) - RentalsHolder 
								End If
				
								CP = ProdSalesHolder 

				
								SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
								SQL = SQL & " FROM CustCatPeriodSales "
								SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
								SQL = SQL & " AND PrimarySalesman = '" & Sls22Find & "'"
								Set rs2 = cnn8.Execute(SQL)
				
								TotSalesCPLY = rs2("CPLYTotalSales")
								TotRentalsCPLY = rs2("CPLYTotalRentals")



								'***********************************
								'***********************************
								' Here is all the ADS variance stuff
								'***********************************
								'***********************************
								LCPADS = Round(rs("TotProductSales"),0)/WorkDaysInLastClosedPeriod
								P3PADS = Round(P3PAvgProductSales,0)/WD_P3PADS
								P12PADS = Round(P12PAvgProductSales,0)/WD_P12PADS
								SPLYADS = Round(SPLYProductsSales,0)/WD_SPLYPADS
								CPADS = Round(CP,0)/WD_CurrentSoFar
								CPLYADS = TotSalesCPLY/WorkDaysInCurrentPLY
				
				
							
								'Sales
								GrandTotalLCPSales = GrandTotalLCPSales + rs("TotProductSales")
								GrandTotal3PAvgSales = GrandTotal3PAvgSales + P3PAvgProductSales 
								GrantTotal12PAvgSales = GrantTotal12PAvgSales + P12PAvgProductSales 
								GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales  
								
								'ADS
								GrandTotalLCPADS = GrandTotalLCPADS + LCPADS 
								GrandTotal3PPADS = GrandTotal3PPADS + P3PADS 
								GrandTotal12PPADS = GrandTotal12PPADS + P12PADS 
								GrandTotalSPLYADS = GrandTotalSPLYADS + SPLYADS 
								GrandTotalCPADS = GrandTotalCPADS + CPADS 
								GrandTotalCPLYADS = GrandTotalCPLYADS + CPLYADS 
								
								'ADS Variance
								GrandTotalLCPvs3PAvgADS = GrandTotalLCPADS - GrandTotal3PPADS 
								GrandTotalLCPvs12PAvg = GrandTotalLCPADS - GrandTotal12PPADS 
								GrandTotalLCPvsSPLY = GrandTotalLCPADS - GrandTotalSPLYADS 
								GrandTotalCPvsvs3PAvgADS = GrandTotalCPADS - GrandTotal3PPADS 
								GrandTotalCPLYvsvs3PAvgADS = GrandTotalCPLYADS - GrandTotal3PPADS 
								
							
							If rs("PrimarySalesman") <> 0 Then
								SQL2 = "INSERT INTO BI_Dashboard ( "
								SQL2 = SQL2 & " PrimarySalesmanNumber "
								SQL2 = SQL2 & " , PrimarySalesmanName "
								SQL2 = SQL2 & " , SalesLCP "
								SQL2 = SQL2 & " , Sales3PPAvg "
								SQL2 = SQL2 & " , Sales12PPAvg "
								SQL2 = SQL2 & " , SalesSPLY "
								SQL2 = SQL2 & " , ADSLCP "
								SQL2 = SQL2 & " , ADS3PPAvg "
								SQL2 = SQL2 & " , ADS12PPAvg "
								SQL2 = SQL2 & " , ADSSPLY "
								SQL2 = SQL2 & " , ADSCP "
								SQL2 = SQL2 & " , ADSCPLY "
								SQL2 = SQL2 & " , VARLCPv3ppAvg "
								SQL2 = SQL2 & " , VARLCPv12PPAvg "
								SQL2 = SQL2 & " , VARLCPvSPLY "
								SQL2 = SQL2 & " , VARCPv3PPAvg "
								SQL2 = SQL2 & " , VARCPvCPLY"
								SQL2 = SQL2 & " , Segment) VALUES ("
								
								SQL2 = SQL2 & "'" & rs("PrimarySalesman") & "'"
								SQL2 = SQL2 & ",'" & GetSalesmanNameBySlsmnSequence(rs("PrimarySalesman")) & "'"
								SQL2 = SQL2 & "," & rs("TotProductSales")
								SQL2 = SQL2 & "," & P3PAvgProductSales
								SQL2 = SQL2 & "," & P12PAvgProductSales
								SQL2 = SQL2 & "," & SPLYProductsSales
								SQL2 = SQL2 & "," & LCPADS 
								SQL2 = SQL2 & "," & P3PADS 
								SQL2 = SQL2 & "," & P12PADS 
								SQL2 = SQL2 & "," & SPLYADS 
								SQL2 = SQL2 & "," & CPADS 
								SQL2 = SQL2 & "," & CPLYADS 
								SQL2 = SQL2 & "," & LCPADS - P3PADS
								SQL2 = SQL2 & "," & LCPADS - P12PADS
								SQL2 = SQL2 & "," & LCPADS - SPLYADS				
								SQL2 = SQL2 & "," & CPADS - P3PADS
								SQL2 = SQL2 & "," & CPADS - CPLYADS 
								SQL2 = SQL2 & ", 'PrimarySalesman')"
								Set rs2 = cnn8.Execute(SQL2)
							End If

					
								rs.movenext
							Loop until rs.eof
						End If



						'***********
						'***********
						' LEFT OVERS
						'***********
						'***********
						If LeftOverSLs1  <> "" Then
						
					      	'This part is a little crazy but now we have to ddo the left over salesman2's
							'Now get all the SPLY Numbers for the leftovers
							SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
							SQL2 = SQL2 & " ,PrimarySalesman "
							SQL2 = SQL2 & " FROM CustCatPeriodSales "
							SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL2 = SQL2 & " AND PrimarySalesman IN (" & LeftOverSLs1 & ") GROUP BY PrimarySalesman "

	
							Set rs = cnn8.Execute(SQL2)
							If Not rs.EOF Then
								Do While Not rs.EOF
								
									SPLYProductsSales = rs2("SPLYTotalSales")
									SPLYTotalRentals = rs2("SPLYTotalRentals")
								
									If SPLYProductsSales <> 0 Then 
								
				
					
									GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales 			
									GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
				
								End If
											
								rs.MoveNext
							Loop
						End IF
					End If
					'***************
					'***************
					' END LEFT OVERS
					'***************
					'***************





'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  P R I M A R Y  S A L E S M A N
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
' C U S T O M E R  T Y P E
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

						SQL = "DELETE FROM BI_Dashboard WHERE Segment = 'CustomerType'"
						Set rs = cnn8.Execute(SQL)


						'Get all Cust Type
					 	
						TotProductSalesCustType = 0
						Tot3PAvgCustType = 0
						TotDollarDiff =0
						TotalNegDiff = 0
						Tot12PAvgCustType = 0
						TotDollarDiff12 = 0
						TotalNegDiff12 = 0
					
						GrandTotalLCPSales = 0 : GrandTotal3PAvgSales = 0 : GrantTotal12PAvgSales = 0 : GrandTotalSPLYSales = 0
						GrandTotalLCPADS = 0 : GrandTotal3PPADS = 0 : GrandTotal12PPADS = 0 : GrandTotalSPLYADS = 0 : GrandTotalCPADS = 0 : GrandTotalCPLYADS = 0
						GrandTotalLCPvs3PAvgADS = 0 : GrandTotalLCPvs12PAvg = 0 : GrandTotalLCPvsSPLY = 0 : GrandTotalCPvsvs3PAvgADS = 0 : GrandTotalCPLYvsvs3PAvgADS = 0
					
						LeftOverCustType = ""

						SQL = "SELECT CustType "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
						SQL = SQL & " GROUP BY CustType "
						SQL = SQL & " EXCEPT "
						SQL = SQL & " SELECT CustType "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
						SQL = SQL & " GROUP BY CustType "
					
						Set rs = cnn8.Execute(SQL)
						If Not rs.EOF Then
							Do While Not rs.EOF
									LeftOverCustType = LeftOverCustType & rs("CustType") & ","
								rs.MoveNext
							Loop
						End IF
					
						If Right(LeftOverCustType,1)="," Then LeftOverCustType = Left(LeftOverCustType,len(LeftOverCustType)-1)
						
						Set rs = cnn8.Execute(SQL)
						
						SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS TotProductSales ,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
						SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
						SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
						SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "	
						SQL = SQL & ",CustCatPeriodSales_ReportData.CustType"
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
						SQL = SQL & " GROUP BY CustCatPeriodSales_ReportData.CustType"
						SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"
					
					
						Set rs = cnn8.Execute(SQL)
						
						If not rs.EOF Then

						ChartElementNumber = 1
						ChartDataCustType = ""
						ChartRemainder = 100
						NextPeriodProj = 0

									
						Do
						
							P3PAvgProductSales = rs("Tot3PPAvg") - rs("Tot3PPAvgRentals")
							P12PAvgProductSales = rs("Tot12PPAvg") - rs("Tot12PPAvgRentals")
							SPLYProductsSales = 0
							SPLYTotalRentals =  0
							LCPADS = 0
							P3PADS = 0
							P12PADS = 0
							SPLYADS = 0
							CPADS = 0
							CPLYADS = 0

							'No link if lcp + 3pp < $1
							TotalToEval = rs("TotProductSales") + P3PAvgProductSales 
							If Not Isnumeric(TotalToEval) Then TotalToEval = 0
			
							Sls22Find = rs("CustType")
							
							'Now get all the SPLY Numbers
							SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
							SQL2 = SQL2 & " FROM CustCatPeriodSales "
							SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL2 = SQL2 & " AND CustType = '" & Sls22Find & "' "
							
							Set rs2 = cnn8.Execute(SQL2)
							If not rs2.EOF Then
								SPLYProductsSales = rs2("SPLYTotalSales")
								SPLYTotalRentals = rs2("SPLYTotalRentals")
							End If

							If IncludeUnposted = True Then
								RentalsHolder = GetCurrentPeriod_PostedRentalsCustType(Sls22Find) + GetCurrentPeriod_UnPostedRentalsCustType(Sls22Find)
								ProdSalesHolder = (GetCurrentPeriod_PostedTotalCustTyp(Sls22Find) + GetCurrentPeriod_UnPostedTotalCustType(Sls22Find)) - RentalsHolder 
							Else
								RentalsHolder = GetCurrentPeriod_PostedRentalsCustType(Sls22Find)
								ProdSalesHolder = GetCurrentPeriod_PostedTotalCustTyp(Sls22Find) - RentalsHolder 
							End If
			
							CP = ProdSalesHolder 


				
							SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
							SQL = SQL & " FROM CustCatPeriodSales "
							SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL = SQL & " AND CustType = '" & Sls22Find & "'"
							Set rs2 = cnn8.Execute(SQL)
							
							TotSalesCPLY = rs2("CPLYTotalSales")
							TotRentalsCPLY = rs2("CPLYTotalRentals")

			
			
							'***********************************
							'***********************************
							' Here is all the ADS variance stuff
							'***********************************
							'***********************************
							LCPADS = Round(rs("TotProductSales"),0)/WorkDaysInLastClosedPeriod
							P3PADS = Round(P3PAvgProductSales,0)/WD_P3PADS
							P12PADS = Round(P12PAvgProductSales,0)/WD_P12PADS
							SPLYADS = Round(SPLYProductsSales,0)/WD_SPLYPADS
							CPADS = Round(CP,0)/WD_CurrentSoFar
							CPLYADS = TotSalesCPLY/WorkDaysInCurrentPLY
				
				
							'Sales
							GrandTotalLCPSales = GrandTotalLCPSales + rs("TotProductSales")
							GrandTotal3PAvgSales = GrandTotal3PAvgSales + P3PAvgProductSales 
							GrantTotal12PAvgSales = GrantTotal12PAvgSales + P12PAvgProductSales 
							GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales  
							
							'ADS
							GrandTotalLCPADS = GrandTotalLCPADS + LCPADS 
							GrandTotal3PPADS = GrandTotal3PPADS + P3PADS 
							GrandTotal12PPADS = GrandTotal12PPADS + P12PADS 
							GrandTotalSPLYADS = GrandTotalSPLYADS + SPLYADS 
							GrandTotalCPADS = GrandTotalCPADS + CPADS 
							GrandTotalCPLYADS = GrandTotalCPLYADS + CPLYADS 
							
							'ADS Variance
							GrandTotalLCPvs3PAvgADS = GrandTotalLCPADS - GrandTotal3PPADS 
							GrandTotalLCPvs12PAvg = GrandTotalLCPADS - GrandTotal12PPADS 
							GrandTotalLCPvsSPLY = GrandTotalLCPADS - GrandTotalSPLYADS 
							GrandTotalCPvsvs3PAvgADS = GrandTotalCPADS - GrandTotal3PPADS 
							GrandTotalCPLYvsvs3PAvgADS = GrandTotalCPLYADS - GrandTotal3PPADS 


							If rs("CustType") <> 0 Then
								SQL2 = "INSERT INTO BI_Dashboard ( "
								SQL2 = SQL2 & " CustomerTypeNumber "
								SQL2 = SQL2 & " , CustomerTypeName "
								SQL2 = SQL2 & " , SalesLCP "
								SQL2 = SQL2 & " , Sales3PPAvg "
								SQL2 = SQL2 & " , Sales12PPAvg "
								SQL2 = SQL2 & " , SalesSPLY "
								SQL2 = SQL2 & " , ADSLCP "
								SQL2 = SQL2 & " , ADS3PPAvg "
								SQL2 = SQL2 & " , ADS12PPAvg "
								SQL2 = SQL2 & " , ADSSPLY "
								SQL2 = SQL2 & " , ADSCP "
								SQL2 = SQL2 & " , ADSCPLY "
								SQL2 = SQL2 & " , VARLCPv3ppAvg "
								SQL2 = SQL2 & " , VARLCPv12PPAvg "
								SQL2 = SQL2 & " , VARLCPvSPLY "
								SQL2 = SQL2 & " , VARCPv3PPAvg "
								SQL2 = SQL2 & " , VARCPvCPLY"
								SQL2 = SQL2 & " , Segment) VALUES ("
								
								SQL2 = SQL2 & "'" & rs("CustType") & "'"
								SQL2 = SQL2 & ",'" & GetCustTypeByCode(rs("CustType")) & "'"
								SQL2 = SQL2 & "," & rs("TotProductSales")
								SQL2 = SQL2 & "," & P3PAvgProductSales
								SQL2 = SQL2 & "," & P12PAvgProductSales
								SQL2 = SQL2 & "," & SPLYProductsSales
								SQL2 = SQL2 & "," & LCPADS 
								SQL2 = SQL2 & "," & P3PADS 
								SQL2 = SQL2 & "," & P12PADS 
								SQL2 = SQL2 & "," & SPLYADS 
								SQL2 = SQL2 & "," & CPADS 
								SQL2 = SQL2 & "," & CPLYADS 
								SQL2 = SQL2 & "," & LCPADS - P3PADS
								SQL2 = SQL2 & "," & LCPADS - P12PADS
								SQL2 = SQL2 & "," & LCPADS - SPLYADS				
								SQL2 = SQL2 & "," & CPADS - P3PADS
								SQL2 = SQL2 & "," & CPADS - CPLYADS 
								SQL2 = SQL2 & ", 'CustomerType')"
								Set rs2 = cnn8.Execute(SQL2)
							End If
								
							rs.movenext
						Loop until rs.eof
					End If



				'***********
				'***********
				' LEFT OVERS
				'***********
				'***********
				If LeftOverTyp  <> "" Then
				
			      	'This part is a little crazy but now we have to ddo the left over salesman2's
					'Now get all the SPLY Numbers for the leftovers
					SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
					SQL2 = SQL2 & " ,SecondarySalesman "
					SQL2 = SQL2 & " FROM CustCatPeriodSales "
					SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
					SQL2 = SQL2 & " AND CustType IN (" & LeftOverTyp & ") GROUP BY CustType"
		
			
					Set rs = cnn8.Execute(SQL2)
					If Not rs.EOF Then
						Do While Not rs.EOF
						
							SPLYProductsSales = rs2("SPLYTotalSales")
							SPLYTotalRentals = rs2("SPLYTotalRentals")
						
					If SPLYProductsSales <> 0 Then 
								
					
						GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales 			
						GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
	
					End If
								
					rs.MoveNext
				Loop
			End IF
		End If
		'***************
		'***************
		' END LEFT OVERS
		'***************
		'***************





'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  C U S T O M E R  T Y P E
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  C U S T O M E R  R E F E R R A L
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

						SQL = "DELETE FROM BI_Dashboard WHERE Segment = 'Referral'"
						Set rs = cnn8.Execute(SQL)


						'Get all Referral Codes
						 	
							TotProductSalesReferral = 0
							Tot3PAvgReferral = 0
							TotDollarDiff =0
							TotalNegDiff = 0
							Tot12PAvgReferral = 0
							TotDollarDiff12 = 0
							TotalNegDiff12 = 0

							GrandTotalLCPSales = 0 : GrandTotal3PAvgSales = 0 : GrantTotal12PAvgSales = 0 : GrandTotalSPLYSales = 0
							GrandTotalLCPADS = 0 : GrandTotal3PPADS = 0 : GrandTotal12PPADS = 0 : GrandTotalSPLYADS = 0 : GrandTotalCPADS = 0 : GrandTotalCPLYADS = 0
							GrandTotalLCPvs3PAvgADS = 0 : GrandTotalLCPvs12PAvg = 0 : GrandTotalLCPvsSPLY = 0 : GrandTotalCPvsvs3PAvgADS = 0 : GrandTotalCPLYvsvs3PAvgADS = 0
						
							LeftOverReferral = ""
						
							SQL = "SELECT ReferralDesc2 "
							SQL = SQL & " FROM CustCatPeriodSales_ReportData "
							SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL = SQL & " GROUP BY ReferralDesc2 "
							SQL = SQL & " EXCEPT "
							SQL = SQL & " SELECT ReferralDesc2 "
							SQL = SQL & " FROM CustCatPeriodSales_ReportData "
							SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
							SQL = SQL & " GROUP BY ReferralDesc2 "
						
						'Response.Write(SQL)
						
							Set rs = cnn8.Execute(SQL)
							If Not rs.EOF Then
								Do While Not rs.EOF
										LeftOverReferral = LeftOverReferral & rs("ReferralDesc2") & ","
									rs.MoveNext
								Loop
							End IF
						
							If Right(LeftOverReferral,1)="," Then LeftOverReferral = Left(LeftOverReferral,len(LeftOverReferral)-1)
							
						'Response.Write(SQL&"<br>")	
							Set rs = cnn8.Execute(SQL)

						SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS TotProductSales ,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
						SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
						SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
						SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
						SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "	
						SQL = SQL & ",ReferralDesc2, Max(Referral) As ReferralCode "
						SQL = SQL & " FROM CustCatPeriodSales_ReportData "
						SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND ReferralDesc2 IS NOT NULL "
						SQL = SQL & " GROUP BY ReferralDesc2 "
						SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"
					
					
						Set rs = cnn8.Execute(SQL)
					
						If not rs.EOF Then

						ChartElementNumber = 1
						ChartDataReferral = ""
						ChartRemainder = 100
						NextPeriodProj = 0
			
												
						Do
						
							P3PAvgProductSales = rs("Tot3PPAvg") - rs("Tot3PPAvgRentals")
							P12PAvgProductSales = rs("Tot12PPAvg") - rs("Tot12PPAvgRentals")
							SPLYProductsSales = 0
							SPLYTotalRentals =  0
							LCPADS = 0
							P3PADS = 0
							P12PADS = 0
							SPLYADS = 0
							CPADS = 0
							CPLYADS = 0
			
							'No link if lcp + 3pp < $1
							TotalToEval = rs("TotProductSales") + P3PAvgProductSales 
							If Not Isnumeric(TotalToEval) Then TotalToEval = 0
			
							Sls22Find = rs("ReferralDesc2")
							
							'Now get all the SPLY Numbers
							SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
							SQL2 = SQL2 & " FROM CustCatPeriodSales "
							SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
							SQL2 = SQL2 & " AND ReferralDesc2 = '" & Sls22Find & "' "
							
							'Response.Write(SQL2)
								
								Set rs2 = cnn8.Execute(SQL2)
								If not rs2.EOF Then
									SPLYProductsSales = rs2("SPLYTotalSales")
									SPLYTotalRentals = rs2("SPLYTotalRentals")
								End If

								If IncludeUnposted = True Then
									RentalsHolder = GetCurrentPeriod_PostedRentalsReferralDesc2(Sls22Find) + GetCurrentPeriod_UnPostedRentalsReferralDesc2(Sls22Find)
									ProdSalesHolder = (GetCurrentPeriod_PostedTotalreferralDesc2(Sls22Find) + GetCurrentPeriod_UnPostedTotalReferralDesc2(Sls22Find)) - RentalsHolder 
								Else
									RentalsHolder = GetCurrentPeriod_PostedRentalsReferralDesc2(Sls22Find)
									ProdSalesHolder = GetCurrentPeriod_PostedTotalreferralDesc2(Sls22Find) - RentalsHolder 
								End If
				
								CP = ProdSalesHolder 

								
								SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
								SQL = SQL & " FROM CustCatPeriodSales "
								SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
								SQL = SQL & " AND ReferralDesc2 = '" & Sls22Find & "'"
				
				
				
								Set rs2 = cnn8.Execute(SQL)
								
								If Not rs2.Eof Then
									TotSalesCPLY = rs2("CPLYTotalSales")
									TotRentalsCPLY = rs2("CPLYTotalRentals")
								Else
									TotSalesCPLY  = 0 
									TotRentalsCPLY = 0
								End IF


								'***********************************
								'***********************************
								' Here is all the ADS variance stuff
								'***********************************
								'***********************************
								LCPADS = Round(rs("TotProductSales"),0)/WorkDaysInLastClosedPeriod
								P3PADS = Round(P3PAvgProductSales,0)/WD_P3PADS
								P12PADS = Round(P12PAvgProductSales,0)/WD_P12PADS
								SPLYADS = Round(SPLYProductsSales,0)/WD_SPLYPADS
								CPADS = Round(CP,0)/WD_CurrentSoFar
								CPLYADS = TotSalesCPLY/WorkDaysInCurrentPLY
				
				
					
								'Sales
								GrandTotalLCPSales = GrandTotalLCPSales + rs("TotProductSales")
								GrandTotal3PAvgSales = GrandTotal3PAvgSales + P3PAvgProductSales 
								GrantTotal12PAvgSales = GrantTotal12PAvgSales + P12PAvgProductSales 
								GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales  
								
								'ADS
								GrandTotalLCPADS = GrandTotalLCPADS + LCPADS 
								GrandTotal3PPADS = GrandTotal3PPADS + P3PADS 
								GrandTotal12PPADS = GrandTotal12PPADS + P12PADS 
								GrandTotalSPLYADS = GrandTotalSPLYADS + SPLYADS 
								GrandTotalCPADS = GrandTotalCPADS + CPADS 
								GrandTotalCPLYADS = GrandTotalCPLYADS + CPLYADS 
								
								'ADS Variance
								GrandTotalLCPvs3PAvgADS = GrandTotalLCPADS - GrandTotal3PPADS 
								GrandTotalLCPvs12PAvg = GrandTotalLCPADS - GrandTotal12PPADS 
								GrandTotalLCPvsSPLY = GrandTotalLCPADS - GrandTotalSPLYADS 
								GrandTotalCPvsvs3PAvgADS = GrandTotalCPADS - GrandTotal3PPADS 
								GrandTotalCPLYvsvs3PAvgADS = GrandTotalCPLYADS - GrandTotal3PPADS 

							
							If rs("ReferralCode") <> 0 Then
								SQL2 = "INSERT INTO BI_Dashboard ( "
								SQL2 = SQL2 & " ReferralCodeNumber "
								SQL2 = SQL2 & " , ReferralCodeDesc2 "
								SQL2 = SQL2 & " , SalesLCP "
								SQL2 = SQL2 & " , Sales3PPAvg "
								SQL2 = SQL2 & " , Sales12PPAvg "
								SQL2 = SQL2 & " , SalesSPLY "
								SQL2 = SQL2 & " , ADSLCP "
								SQL2 = SQL2 & " , ADS3PPAvg "
								SQL2 = SQL2 & " , ADS12PPAvg "
								SQL2 = SQL2 & " , ADSSPLY "
								SQL2 = SQL2 & " , ADSCP "
								SQL2 = SQL2 & " , ADSCPLY "
								SQL2 = SQL2 & " , VARLCPv3ppAvg "
								SQL2 = SQL2 & " , VARLCPv12PPAvg "
								SQL2 = SQL2 & " , VARLCPvSPLY "
								SQL2 = SQL2 & " , VARCPv3PPAvg "
								SQL2 = SQL2 & " , VARCPvCPLY"
								SQL2 = SQL2 & " , Segment) VALUES ("
								
								SQL2 = SQL2 & "'" & rs("ReferralCode") & "'"
								SQL2 = SQL2 & ",'" & rs("ReferralDesc2") & "'"
								SQL2 = SQL2 & "," & rs("TotProductSales")
								SQL2 = SQL2 & "," & P3PAvgProductSales
								SQL2 = SQL2 & "," & P12PAvgProductSales
								SQL2 = SQL2 & "," & SPLYProductsSales
								SQL2 = SQL2 & "," & LCPADS 
								SQL2 = SQL2 & "," & P3PADS 
								SQL2 = SQL2 & "," & P12PADS 
								SQL2 = SQL2 & "," & SPLYADS 
								SQL2 = SQL2 & "," & CPADS 
								SQL2 = SQL2 & "," & CPLYADS 
								SQL2 = SQL2 & "," & LCPADS - P3PADS
								SQL2 = SQL2 & "," & LCPADS - P12PADS
								SQL2 = SQL2 & "," & LCPADS - SPLYADS				
								SQL2 = SQL2 & "," & CPADS - P3PADS
								SQL2 = SQL2 & "," & CPADS - CPLYADS 
								SQL2 = SQL2 & ", 'Referral')"
								Set rs2 = cnn8.Execute(SQL2)
							End If
							
								rs.movenext
							Loop until rs.eof
						End If



						'***********
						'***********
						' LEFT OVERS
						'***********
						'***********
						If LeftOverSLs2  <> "" Then
						
								SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
								SQL2 = SQL2 & " FROM CustCatPeriodSales "
								SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
								SQL2 = SQL2 & " AND ReferralDesc2 = '" & Sls22Find & "' "
				
					
							Set rs = cnn8.Execute(SQL2)
							If Not rs.EOF Then
								Do While Not rs.EOF
								
									SPLYProductsSales = rs2("SPLYTotalSales")
									SPLYTotalRentals = rs2("SPLYTotalRentals")
								
									If SPLYProductsSales <> 0 Then 
												
										Response.Write("<tr>")
										
										Response.Write("<td align='left' class='smaller-detail-line'>" & rs("ReferralCode") & " - " & rs("ReferralDesc2") & "</td>")									
										
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						
						
										Response.Write("<td  class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(SPLYProductsSales,0,-2,0) & "</td>")
				
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")						
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
										Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
				
									
										If SPLYProductsSales * -1 > 0 Then
											Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(SPLYProductsSales * -1,0,-2,0) & "</td>")				
										Else
											Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(SPLYProductsSales * -1,0,-2,0) & "</td>")							
										End If
										
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
										Response.Write("<td class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(0,0,-2,0) & "</td>")
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
										Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
								
										Response.Write("</tr>")
									
										GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales 			
										GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
					
									End If
												
									rs.MoveNext
								Loop
							End IF
						End If
						'***************
						'***************
						' END LEFT OVERS
						'***************
						'***************





'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  C U S T O M E R  R E F E R R A L
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'S E G M E N T  T A B  D A T A 
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'Runs through every customer in the _ReportData table & calculates everything needed for the tab data

	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open MUV_READ("ClientCnnString") 
	
	SQL = "DELETE FROM BI_DashboardSegmentTabs1"
	SQL = "DELETE FROM BI_DashboardSegmentTabs"
	Set rs = cnn8.Execute(SQL)

	FilterSalesDollars = 100
	FilterPercentage = 10

	For z = 1 to 10'10 ' There are 10 tabs
	
		Select Case z
		
			Case 1 ' UP tab
			
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " AND  LCPTotSalesAllCats > (Total3PPSalesAllCats /3) "
				SQL = SQL & " ORDER BY CustCatPeriodSales_ReportData.CustNum"
				
				TabName = "UP"

			Case 2 ' Down tab
	
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " AND  LCPTotSalesAllCats <= (Total3PPSalesAllCats /3) "
				SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
				SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
				SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
				SQL = SQL & " ORDER BY CustCatPeriodSales_ReportData.CustNum"

				TabName = "DOWN"

			Case 3 ' $0 Sales tab
	
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " AND  LCPTotSalesAllCats <= 0 "
				SQL = SQL & " ORDER BY CustCatPeriodSales_ReportData.CustNum"

				TabName = "ZEROSALES"
				
			Case 4 ' Ruled Out
	
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " AND  LCPTotSalesAllCats <= (Total3PPSalesAllCats /3) "
				SQL = SQL & " ORDER BY CustCatPeriodSales_ReportData.CustNum"
				
				TabName = "RULEDOUT"
				
			Case 5 ' all
	
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " ORDER BY CustCatPeriodSales_ReportData.CustNum"

				TabName = "ALL"

			Case 6 ' top50
	
				REM - Don't do anything for top 50, uses the ALL records
			
			Case 7 ' bottom50
	
				REM - Don't do anything for bottom 50, uses the ALL records
				
			Case 8 ' mcs
	
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "	
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				SQL = SQL & " AND  AR_Customer.MonthlyContractedSalesDollars IS NOT NULL"

				TabName = "MCS"

			Case 9 ' high roi
			
				SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
				SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
				SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 

				TabName = "HIGHROI"

			Case 10 ' cat breakdown
	
				REM - Don't do anything for category breakdown
			
			End Select
		
		If z <> 6 and z <> 7 and z <> 10 Then
		
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'HERE IS THE MAIN BUILD LOGIC
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.ConnectionTimeout = 120
				cnn8.open (Session("ClientCnnString"))
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3
				Set rs = cnn8.Execute(SQL)
				
			
					Do While Not rs.EOF
			
						ShowThisRecord = True
			
							
						If ShowThisRecord <> False Then			
						
							PrimarySalesMan =  ""
							SecondarySalesMan =  ""
							CustomerType =  ""
							CustTypeCode = ""
							CustomerTypeDesc ="" 
							ReferralCodeForSQL  =""
							SelectedCustomerID = rs("CustNum")
							CustName = GetCustNameByCustNum(SelectedCustomerID)	
							If CustName <> "" Then CustName = Replace(CustName,"'","")
							
							'Extra Fields for Filtering
							SQL4 = "SELECT * FROM AR_Customer WHERE CustNum = '" & SelectedCustomerID & "'"
							Set rs4 = Server.CreateObject("ADODB.Recordset")
							rs4.CursorLocation = 3
							Set rs4= cnn8.Execute(SQL4 )
			
							If Not rs4.Eof Then
								PrimarySalesMan = rs4("Salesman")
								SecondarySalesMan = rs4("SecondarySalesman")
								ReferralCode = rs4("ReferalCode")
								ReferralCodeForSQL  = rs4("ReferalCode")
								CustTypeCode = rs4("CustType")
								CustomerTypeDesc = GetCustTypeByCode(rs4("CustType"))
								CustomerTypeDesc = Replace(CustomerTypeDesc ,"CLIENT","")
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
							
							If TotalEquipmentValue > 0 Then	
								'LCPGP = LCPSales - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,SelectedCustomerID)
								LCPGP = LCPSales - rs("TotalCostAllCats")
								ThreePAvgGP = ThreePPAvgSales - ( TotalCostByPeriodSeqPrior3P(PeriodSeqBeingEvaluated,SelectedCustomerID) / 3 )
								If LCPGP <> 0 Then ROI = TotalEquipmentValue/LCPGP Else ROI = ""
								If ThreePAvgGP <> 0 Then ROI3P = TotalEquipmentValue/ThreePAvgGP Else ROI3P = ""
								'Calculate 12P ROI
								TwelvePAvgGP = TwelvePPAvgSales -- ( TotalCostByPeriodSeqPrior12P(PeriodSeqBeingEvaluated,SelectedCustomerID) / 12 )
								If TwelvePAvgGP <> 0 Then ROI12P = TotalEquipmentValue/TwelvePAvgGP Else ROI12P = ""
							End If
			
							If Not IsNumeric(ThreePPAvgSales) Then ThreePPAvgSales = 0
							If Not IsNumeric(TwelvePPAvgSales) Then TwelvePPAvgSales = 0
			
			
							If TabName = "DOWN" Then
							
										' HERE ARE THE RULES
										' THE FOUR RULES (5 NOW)
						
										'1. If Current (adjusted for days) => 3Pavg or 12Pavg - Dont show
										If CurrentPSales >= ThreePPAvgSales OR CurrentPSales >= TwelvePPAvgSales Then  'If the current is already greater or equal, we don't need to adjust for days, we're already ok
											ShowThisRecord = False
										Else  'We need to adjust for days & fiure it out
											ForecastedCurrent = (CurrentPSales / WorkDaysSoFar) * WorkDaysInCurrentPeriod 
											If ForecastedCurrent >= ThreePPAvgSales OR ForecastedCurrent >= TwelvePPAvgSales Then
												ShowThisRecord = False
											End If
										End If
											
										'2. If LCP => 12pAVG - Dont Show
										If LCPSales >= TwelvePPAvgSales Then
											ShowThisRecord = False
										End If
										
										'3. If LCP => SPLY - Dont Show
										If LCPSales >= SamePLYSales Then
											ShowThisRecord = False
										End If
										
										'4 The NEW 4th Rule
										If cdbl(LCPSales + PP1Sales  + PP2Sales) / 3  >= cdbl(ThreePPAvgSales) Then
											ShowThisRecord = False
										End If
						
										'5 The New 5th Rule
										If cdbl(LCPSales + PP1Sales  + PP2Sales) / 3  >= cdbl(TwelvePPAvgSales) Then
											ShowThisRecord = False
										End If
										
										'6. If 3PROI > 10 - Override anything else and Show
										If Not Isnull(ROI3P) Then
											If IsNumeric(ROI3P) Then
												If ROI3P > 10 Then
													ShowThisRecord = True
												End IF
											End If
										End If
							End If
							
							If TabName = "RULEDOUT" Then

								ShowThisRecord = False ' Interesting, these we set to false
								RulesApplied = ""
								
										' HERE ARE THE RULES
										' THE FOUR RULES (5 NOW)
						
										'1. If Current (adjusted for days) => 3Pavg or 12Pavg - Dont show
										If CurrentPSales >= ThreePPAvgSales OR CurrentPSales >= TwelvePPAvgSales Then  'If the current is already greater or equal, we don't need to adjust for days, we're already ok
											ShowThisRecord = True
										Else  'We need to adjust for days & fiure it out
											ForecastedCurrent = (CurrentPSales / WorkDaysSoFar) * WorkDaysInCurrentPeriod 
											If ForecastedCurrent >= ThreePPAvgSales OR ForecastedCurrent >= TwelvePPAvgSales Then
												ShowThisRecord = True
											End If
										End If
											
										'2. If LCP => 12pAVG - Dont Show
										If LCPSales >= TwelvePPAvgSales Then
											ShowThisRecord = True
										End If
										
										'3. If LCP => SPLY - Dont Show
										If LCPSales >= SamePLYSales Then
											ShowThisRecord = True
										End If
										
										'4 The NEW 4th Rule
										If cdbl(LCPSales + PP1Sales  + PP2Sales) / 3  >= cdbl(ThreePPAvgSales) Then
											ShowThisRecord = True
										End If
						
										'5 The New 5th Rule
										If cdbl(LCPSales + PP1Sales  + PP2Sales) / 3  >= cdbl(TwelvePPAvgSales) Then
											ShowThisRecord = True
										End If
										
										'6. If 3PROI > 10 - Override anything else and Show
										If Not Isnull(ROI3P) Then
											If IsNumeric(ROI3P) Then
												If ROI3P > 10 Then
													ShowThisRecord = False 
												End IF
											End If
										End If

							
							End IF

							If TabName = "HIGHROI" Then
							
								If TotalEquipmentValue > 0 Then	
									If ROI3P <> "" And ROI <> "" Then
										If ROI3P < 10 And ROI < 10 Then ShowThisRecord = False
									Else
										If ROI12P <> "" Then
											If ROI12P < 10 Then ShowThisRecord = False
										End If
									End If
								Else
									ShowThisRecord = False
								End If	
								
							End If
							
							If ShowThisRecord <> False Then
							
			
								If NOT IsNumeric(SecondarySalesMan) Then SecondarySalesMan = 0		
								If NOT IsNumeric(ReferralCode) Then ReferralCode  = 0		
								If NOT IsNumeric(CustTypeCode) Then CustTypeCode = 0	
								If NOT IsNumeric(CustTypeCode) Then CustTypeCode = 0
								
											SQL2 = "INSERT INTO BI_DashboardSegmentTabs1( "
											SQL2 = "INSERT INTO BI_DashboardSegmentTabs( "
											SQL2 = SQL2 & "  TAB "
											SQL2 = SQL2 & " , CustID "
											SQL2 = SQL2 & " , CustName "
											SQL2 = SQL2 & " , LCPv3PAvg "
											SQL2 = SQL2 & " , DayImpact "
											SQL2 = SQL2 & " , ADS "
											SQL2 = SQL2 & " , LCPv12PAvg "
											SQL2 = SQL2 & " , PP1Sales "
											SQL2 = SQL2 & " , PP2Sales "
											SQL2 = SQL2 & " , LCPSales "
											SQL2 = SQL2 & " , ThreePAvgSales "
											SQL2 = SQL2 & " , TwelvePAvgSales "
											SQL2 = SQL2 & " , CPSales "
											SQL2 = SQL2 & " , SPLYSales "
											SQL2 = SQL2 & " , MCS "
											SQL2 = SQL2 & " , LCPvMCS "
											SQL2 = SQL2 & " , ThreePAvgvMCS "								
											SQL2 = SQL2 & " , TwelvePAvgvMCS "								
											SQL2 = SQL2 & " , CPvMCS "								
											SQL2 = SQL2 & " , LCPROI "								
											SQL2 = SQL2 & " , ThreePAvgROI "	
											SQL2 = SQL2 & " , TwelvePAvgROI "							
											SQL2 = SQL2 & " , EqpValue "								
											SQL2 = SQL2 & " , PrimarySalesmanNumber "								
											SQL2 = SQL2 & " , PrimarySalesmanName "		
											SQL2 = SQL2 & " , SecondarySalesmanNumber "																
											SQL2 = SQL2 & " , SecondarySalesmanName "																
											SQL2 = SQL2 & " , CustomerTypeNumber "																
											SQL2 = SQL2 & " , CustomerTypeName "																
											SQL2 = SQL2 & " , ReferralCode "																
											SQL2 = SQL2 & " , ReferralDesc2 "																
											SQL2 = SQL2 & " ) VALUES ("
											
											SQL2 = SQL2 & "'" & tabName & "'"
											SQL2 = SQL2 & ",'" & SelectedCustomerID & "'"
											SQL2 = SQL2 & ",'" & CustName & "'"
											SQL2 = SQL2 & "," & LCPvs3PAvgSales
											SQL2 = SQL2 & "," & DayImpact
											SQL2 = SQL2 & "," & ADS_LastClosed
											SQL2 = SQL2 & "," & LCPvs12PAvgSales 
											SQL2 = SQL2 & "," & PP1Sales 
											SQL2 = SQL2 & "," & PP2Sales 
											SQL2 = SQL2 & "," & LCPSales 
											SQL2 = SQL2 & "," & ThreePPAvgSales
											SQL2 = SQL2 & "," & TwelvePPAvgSales
											SQL2 = SQL2 & "," & CurrentPSales 
											SQL2 = SQL2 & "," & SamePLYSales 
											If MonthlyContractedSalesDollars <> "" Then
												SQL2 = SQL2 & ",'" & MonthlyContractedSalesDollars & "'"
												SQL2 = SQL2 & ",'" & LCPSales-MonthlyContractedSalesDollars& "'"
												SQL2 = SQL2 & ",'" & ThreePPAvgSales-MonthlyContractedSalesDollars& "'"
												SQL2 = SQL2 & ",'" & TwelvePPAvgSales-MonthlyContractedSalesDollars& "'"
												SQL2 = SQL2 & ",'" & CurrentPSales-MonthlyContractedSalesDollars& "'"
											Else
												SQL2 = SQL2 & ",NULL"
												SQL2 = SQL2 & ",NULL"
												SQL2 = SQL2 & ",NULL"
												SQL2 = SQL2 & ",NULL"
												SQL2 = SQL2 & ",NULL"
											End If
											SQL2 = SQL2 & ",'" & ROI  & "'"
											SQL2 = SQL2 & ",'" & ROI3P  & "'"
											SQL2 = SQL2 & ",'" & ROI12P  & "'"
											SQL2 = SQL2 & "," & TotalEquipmentValue 
											SQL2 = SQL2 & "," & PrimarySalesMan 
											SQL2 = SQL2 & ",'" &  GetSalesmanNameBySlsmnSequence(PrimarySalesMan) & "'"
											SQL2 = SQL2 & "," & SecondarySalesMan 
											SQL2 = SQL2 & ",'" & GetSalesmanNameBySlsmnSequence(SecondarySalesMan) & "'"
											SQL2 = SQL2 & "," & CustTypeCode 
											SQL2 = SQL2 & ",'" & CustomerTypeDesc & "'"
											SQL2 = SQL2 & "," & ReferralCodeForSQL  
											SQL2 = SQL2 & ",'" & GetCustRefDesc2ByReferralCode(ReferralCodeForSQL) & "'"
											SQL2 = SQL2 & ")"
											
											Response.Write("<br><br>" & SQL2 & "<br><br>")
											
											Set rs2 = cnn8.Execute(SQL2)
			

										
										end if
										end if
										
										
										rs.movenext
										
										Loop
					
				End If ' for the 6 & 7
				
	
	Next

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'E O F  S E G M E N T  T A B  D A T A 
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
						
	
										
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
		End If	
		
	Else ' is the biz in tel  module enabled
	
		Call SetClientCnnString
				
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
			
		'Get next Entry Thread for use in the SC_AuditLogDLaunch table
		On Error Goto 0
		Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
		cnnAuditLog.open MUV_READ("ClientCnnString") 
		Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
		rsAuditLog.CursorLocation = 3 
		Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
		If Not rsAuditLog.EOF Then 
		If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
		Else
		EntryThread = 1
		End If
		set rsAuditLog = nothing
		cnnAuditLog.close
		set cnnAuditLog = nothing


		WriteResponse ("Skipping the client " & ClientKey & " because the Biz Intel module is not enabled.<BR>")
		
	End If ' is the Service  module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")

'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub


Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'MCS Rebuild Helper'"
	SQL = SQL & ",'/directlaunch/bizintel/mcs_rebuild_helper_launch.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	'Response.write("<BR>" & SQL & "<BR>")
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open Session("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub


Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function


'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************


%>