<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../../inc/InSightFuncs_AR_AP.asp"-->


<%
	Segment = Request.QueryString("p")
	
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
	
			SQL ="SELECT CategoryNameGetTerm, Category, SUM(TotalSales) AS TotalSales, SUM([3PriorPeriodsAeverage]) AS ThreePriorPeriodsAeverage, SUM(ThisPeriodLastYearSales) AS ThisPeriodLastYearSales "
			SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
			SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg"
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " WHERE SecondarySalesman = " & Segment & " AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND Category <> 0 "
		    SQL = SQL & " GROUP BY CategoryNameGetTerm, Category"

		Case "Primary"
	
			SQL ="SELECT CategoryNameGetTerm, Category, SUM(TotalSales) AS TotalSales, SUM([3PriorPeriodsAeverage]) AS ThreePriorPeriodsAeverage, SUM(ThisPeriodLastYearSales) AS ThisPeriodLastYearSales "
			SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
			SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg"
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " WHERE PrimarySalesman = " & Segment & " AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND Category <> 0 "
		    SQL = SQL & " GROUP BY CategoryNameGetTerm, Category"

		Case "CustType"
	
			SQL ="SELECT CategoryNameGetTerm, Category, SUM(TotalSales) AS TotalSales, SUM([3PriorPeriodsAeverage]) AS ThreePriorPeriodsAeverage, SUM(ThisPeriodLastYearSales) AS ThisPeriodLastYearSales "
			SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
			SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg"
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " WHERE PrimarySalesman = " & Segment & " AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND Category <> 0 "
		    SQL = SQL & " GROUP BY CategoryNameGetTerm, Category"
		    
	End Select	
	
'	Response.write(SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.ConnectionTimeout = 120
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)
			
		Do While Not rs.EOF
		
			CategoryNumber = rs("Category")
			Category = rs("Category") & " - " & GetCategoryByID(rs("Category"))
			TotalSales = rs("TotalSales")
			ThreePriorPeriodsAeverage = rs("ThreePriorPeriodsAeverage")							
			ThisPeriodLastYearSales = rs("ThisPeriodLastYearSales")
			TwelvePPSales = rs("Tot12PPAvg")

			CurrentPSales = GetCurrent_PostedTotal_ByCatBySecondary(PeriodSeqBeingEvaluated,CategoryNumber,Segment) ' + GetCurrent_UnPostedTotal_ByCatBySecondary(PeriodSeqBeingEvaluated,CategoryNumber,Segment)
							
			IF LEN(JSON)>0 Then
				JSON=JSON+","
			END If
			JSON=JSON+"{"
			JSON=JSON & """Category"":""" & Category & """"
			JSON=JSON+","
			
			JSON=JSON & """TotalSales"":""" & FormatCurrency(TotalSales,0) & """"
			JSON=JSON+","
			JSON=JSON & """CurrentSales"":""" & FormatCurrency(CurrentPSales,0) & """"
			JSON=JSON+","
			JSON=JSON & """ThreePriorPeriodsAeverage"":""" & FormatCurrency(ThreePriorPeriodsAeverage,0) & """"
			JSON=JSON+","
			JSON=JSON & """TwelvePriorPeriodsAeverage"":""" & FormatCurrency(TwelvePPSales,0) & """"
			JSON=JSON+","
			JSON=JSON & """ThisPeriodLastYearSales"":""" & FormatCurrency(ThisPeriodLastYearSales,0) & """"


			JSON=JSON+","
			JSON=JSON & """CurrentSalesVariance"":""" & FormatCurrency((CurrentPSales-TotalSales),0) & """"
			JSON=JSON+","
			JSON=JSON & """ThreePriorPeriodsAeverageVariance"":""" & FormatCurrency((ThreePriorPeriodsAeverage-TotalSales),0) & """"
			JSON=JSON+","
			JSON=JSON & """TwelvePriorPeriodsAeverageVariance"":""" & FormatCurrency((TwelvePPSales-TotalSales),0) & """"
			JSON=JSON+","
			JSON=JSON & """ThisPeriodLastYearSalesVariance"":""" & FormatCurrency((ThisPeriodLastYearSales-TotalSales),0) & """"
			

            JSON=JSON & "}"
		    'Response.Write("</tr>")
			
		rs.movenext
				
	Loop

	'retData="{""orderby"":""" & orderValue & """,""draw"": " & CLng(Request.QueryString("draw")) & ",""recordsTotal"": " & nRecordCount & ",""recordsFiltered"": " & nRecordCount & ",""data"": [" & JSONdata & "],""byRegionData"":"+GetQtyCustByRegion()+"}"
	JSON="{""data"":[" & JSON & "]}"
		
	Response.AddHeader "Content-Type", "application/json"
	response.write JSON


Function GetCurrent_PostedTotal_ByCat(passedPeriodBeingEvaluated, passedCategory)


	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_PostedTotal_ByCat = 0

	Set cnnGetCurrent_PostedTotal_ByCat = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCat.open Session("ClientCnnString")

	SQLGetCurrent_PostedTotal_ByCat = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByCat = SQLGetCurrent_PostedTotal_ByCat & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1
	SQLGetCurrent_PostedTotal_ByCat = SQLGetCurrent_PostedTotal_ByCat & " AND CategoryID = " & passedCategory
	
	Set rsGetCurrent_PostedTotal_ByCat = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCat.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCat = cnnGetCurrent_PostedTotal_ByCat.Execute(SQLGetCurrent_PostedTotal_ByCat)

	If not rsGetCurrent_PostedTotal_ByCat.EOF Then resultGetCurrent_PostedTotal_ByCat = rsGetCurrent_PostedTotal_ByCat("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCat) Then resultGetCurrent_PostedTotal_ByCat = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCat.Close
	set rsGetCurrent_PostedTotal_ByCat= Nothing
	cnnGetCurrent_PostedTotal_ByCat.Close	
	set cnnGetCurrent_PostedTotal_ByCat= Nothing

	
	GetCurrent_PostedTotal_ByCat = resultGetCurrent_PostedTotal_ByCat

End Function


Function GetCurrent_UnpostedTotal_ByCat(passedPeriodBeingEvaluated,passedCategory)

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_UnpostedTotal_ByCat = 0

	Set cnnGetCurrent_UnpostedTotal_ByCat = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCat.open Session("ClientCnnString")

	SQLGetCurrent_UnpostedTotal_ByCat = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnpostedTotal_ByCat = SQLGetCurrent_UnpostedTotal_ByCat & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1
	SQLGetCurrent_UnpostedTotal_ByCat = SQLGetCurrent_UnpostedTotal_ByCat & " AND CategoryID = " & passedCategory

	Set rsGetCurrent_UnpostedTotal_ByCat = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCat.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCat = cnnGetCurrent_UnpostedTotal_ByCat.Execute(SQLGetCurrent_UnpostedTotal_ByCat)

	If not rsGetCurrent_UnpostedTotal_ByCat.EOF Then resultGetCurrent_UnpostedTotal_ByCat = rsGetCurrent_UnpostedTotal_ByCat("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCat) Then resultGetCurrent_UnpostedTotal_ByCat = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCat.Close
	set rsGetCurrent_UnpostedTotal_ByCat= Nothing
	cnnGetCurrent_UnpostedTotal_ByCat.Close	
	set cnnGetCurrent_UnpostedTotal_ByCat= Nothing
	
	GetCurrent_UnpostedTotal_ByCat = resultGetCurrent_UnpostedTotal_ByCat 

End Function

Function GetCurrent_PostedTotal_ByCatBySecondary(passedPeriodBeingEvaluated, passedCategory, passedSecondary)


	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_PostedTotal_ByCatBySecondary = 0

	Set cnnGetCurrent_PostedTotal_ByCatBySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCatBySecondary.open Session("ClientCnnString")

	SQLGetCurrent_PostedTotal_ByCatBySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByCatBySecondary = SQLGetCurrent_PostedTotal_ByCatBySecondary & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1
	SQLGetCurrent_PostedTotal_ByCatBySecondary = SQLGetCurrent_PostedTotal_ByCatBySecondary & " AND CategoryID = " & passedCategory
	SQLGetCurrent_PostedTotal_ByCatBySecondary = SQLGetCurrent_PostedTotal_ByCatBySecondary & " AND SecondarySalesman = " & passedSecondary
	
	Set rsGetCurrent_PostedTotal_ByCatBySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCatBySecondary.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCatBySecondary = cnnGetCurrent_PostedTotal_ByCatBySecondary.Execute(SQLGetCurrent_PostedTotal_ByCatBySecondary)

	If not rsGetCurrent_PostedTotal_ByCatBySecondary.EOF Then resultGetCurrent_PostedTotal_ByCatBySecondary = rsGetCurrent_PostedTotal_ByCatBySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCatBySecondary) Then resultGetCurrent_PostedTotal_ByCatBySecondary = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCatBySecondary.Close
	set rsGetCurrent_PostedTotal_ByCatBySecondary= Nothing
	cnnGetCurrent_PostedTotal_ByCatBySecondary.Close	
	set cnnGetCurrent_PostedTotal_ByCatBySecondary= Nothing

	
	GetCurrent_PostedTotal_ByCatBySecondary = resultGetCurrent_PostedTotal_ByCatBySecondary

End Function


Function GetCurrent_UnpostedTotal_ByCatBySecondary(passedPeriodBeingEvaluated,passedCategory,passedSecondary)

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_UnpostedTotal_ByCatBySecondary = 0

	Set cnnGetCurrent_UnpostedTotal_ByCatBySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCatBySecondary.open Session("ClientCnnString")

	SQLGetCurrent_UnpostedTotal_ByCatBySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnpostedTotal_ByCatBySecondary = SQLGetCurrent_UnpostedTotal_ByCatBySecondary & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1
	SQLGetCurrent_UnpostedTotal_ByCatBySecondary = SQLGetCurrent_UnpostedTotal_ByCatBySecondary & " AND CategoryID = " & passedCategory
	SQLGetCurrent_UnpostedTotal_ByCatBySecondary = SQLGetCurrent_UnpostedTotal_ByCatBySecondary & " AND SecondarySalesman = " & passedSecondary	

	Set rsGetCurrent_UnpostedTotal_ByCatBySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCatBySecondary.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCatBySecondary = cnnGetCurrent_UnpostedTotal_ByCatBySecondary.Execute(SQLGetCurrent_UnpostedTotal_ByCatBySecondary)

	If not rsGetCurrent_UnpostedTotal_ByCatBySecondary.EOF Then resultGetCurrent_UnpostedTotal_ByCatBySecondary = rsGetCurrent_UnpostedTotal_ByCatBySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCatBySecondary) Then resultGetCurrent_UnpostedTotal_ByCatBySecondary = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCatBySecondary.Close
	set rsGetCurrent_UnpostedTotal_ByCatBySecondary= Nothing
	cnnGetCurrent_UnpostedTotal_ByCatBySecondary.Close	
	set cnnGetCurrent_UnpostedTotal_ByCatBySecondary= Nothing
	
	GetCurrent_UnpostedTotal_ByCatBySecondary = resultGetCurrent_UnpostedTotal_ByCatBySecondary 

End Function

%>

