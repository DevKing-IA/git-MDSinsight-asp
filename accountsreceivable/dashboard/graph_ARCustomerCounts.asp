<%


'**********************************************************************************
'HELPER FUNCTION TO GET LAST DAY OF MONTH
'**********************************************************************************
Function GetLastDayofMonth(aDate)
    dim intMonth
    dim dteFirstDayNextMonth

    dtefirstdaynextmonth = dateserial(year(adate),month(adate) + 1, 1)
    GetLastDayofMonth = Day(DateAdd ("d", -1, dteFirstDayNextMonth))
End Function

'**********************************************************************************


Set cnnARCustomerCountGraph = Server.CreateObject("ADODB.Connection")
cnnARCustomerCountGraph.open (Session("ClientCnnString"))
Set rsARCustomerCountGraph = Server.CreateObject("ADODB.Recordset")
rsARCustomerCountGraph.CursorLocation = 3

%>
<div id="chartdivARCustomerCounts" style="width: 100%; height: 400px; margin: 0 auto"></div>	
<%

'Now Build the Chart Data Variable

amChartDataARCustCounts = ""

currentMonthInYear = Month(Now())
currentYear = Year(Now())

If cInt(currentMonthInYear) = 12 Then
	
	For i = 1 to 12 'Loop through all months of this year
	
		loopMonth = i
		loopYear = currentYear

		%><!--#include file="graph_ARCustomerCountsBuildData.asp"--><%
	Next
	
Else
	
	
	For i = cInt(currentMonthInYear) to 12 'Loop through the months in the previous year
	
		loopMonth = i
		loopYear = currentYear - 1
		
		%><!--#include file="graph_ARCustomerCountsBuildData.asp"--><%		        
	Next

	For i = 1 to cInt(currentMonthInYear) 'Loop through all the remaining months of this year

		loopMonth = i
		loopYear = currentYear
		
		%><!--#include file="graph_ARCustomerCountsBuildData.asp"--><% 
		        
	Next
	

End If

'Strip last comma

If amChartDataARCustCounts <> "" Then
	amChartDataARCustCounts = Left(amChartDataARCustCounts, len(amChartDataARCustCounts)-1)
End If


cnnARCustomerCountGraph.Close
Set cnnARCustomerCountGraph = Nothing
Set rsARCustomerCountGraph = Nothing
%>